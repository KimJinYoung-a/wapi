<%
function getOutmallRefOrgOrderNO(iorderno,iorderdtlsn,isellsite)
    Dim strSQL
    getOutmallRefOrgOrderNO = iorderno

    strSQL = "select ref_outmallorderserial"
    strSQL = strSQL & " from db_temp.dbo.tbl_XSite_tmporder WITH(NOLOCK)"
    strSQL = strSQL & " where outmallorderserial='"&iorderno&"'"
    strSQL = strSQL & " and orgDetailkey='"&iorderdtlsn&"'"
    strSQL = strSQL & " and sellsite='"&isellsite&"' "
    rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    IF (NOT rsget.Eof) then
        getOutmallRefOrgOrderNO = rsget("ref_outmallorderserial")
    end if
    rsget.Close
end function

'옵션추가금액 새로 만든 Que Table에 저장
Function SugiOptionQueLogInsert(imallid, iapiaction, imidx, iresultcode, ilastErrMsg, ilastUpdateid)
	Dim strSQL

	ilastErrMsg = replace(ilastErrMsg, "'", "′")
	strSQL = ""
	strSQL = strSQL & " INSERT INTO [db_etcmall].[dbo].[tbl_outmall_Option_API_Que] (mallid, apiAction, midx, priority, regdate, readdate, findate, resultCode, lastErrMsg, lastUserid) VAlUES " & VBCRLF
	strSQL = strSQL & " ('"& imallid &"', '"& iapiaction &"', '"& imidx &"', '999999', getdate(), getdate(), getdate(), '"& iresultcode &"', '"& LEFT(ilastErrMsg, 100) &"', '"& ilastUpdateid &"') " & VBCRLF
	dbget.Execute strSQL
	If iresultcode = "OK" Then
		If (imallid = "gsshop") Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
				If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "PRICE") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.[tbl_outmall_Option_API_Que] " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and midx = '"&imidx&"' " & VBCRLF
					If iapiaction = "PRICE" Then
						strSQL = strSQL & " and apiAction in ('PRICE') " & VBCRLF
					ElseIf iapiaction = "EditSellYn" Then
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
					Else
						strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		End If
	End If
End Function

'각 옵션별 새로 만든 Que Table에 저장
Function SugiQueLogInsertByOption(imallid, iapiaction, iitemid, iitemoption, iresultcode, ilastErrMsg, ilastUpdateid)
	Dim strSQL

	ilastErrMsg = replace(ilastErrMsg, "'", "′")
	strSQL = ""
	strSQL = strSQL & " INSERT INTO [db_etcmall].[dbo].[tbl_outmall_newItem_API_Que] (mallid, apiAction, itemid, itemoption, priority, regdate, readdate, findate, resultCode, lastErrMsg, lastUserid) VAlUES " & VBCRLF
	strSQL = strSQL & " ('"& imallid &"', '"& iapiaction &"', '"& iitemid &"', '"&itemoption&"', '999999', getdate(), getdate(), getdate(), '"& iresultcode &"', '"& LEFT(ilastErrMsg, 100) &"', '"& ilastUpdateid &"') " & VBCRLF
	dbget.Execute strSQL
	If iresultcode = "OK" Then
		If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
			If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "PRICE") Then
				strSQL = ""
				strSQL = strSQL & " UPDATE db_etcmall.dbo.[tbl_outmall_newItem_API_Que] " & VBCRLF
				strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
				strSQL = strSQL & " ,findate=getdate() " & VBCRLF
				strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
				strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
				strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
				strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
				strSQL = strSQL & " and itemoption = '"&iitemoption&"' " & VBCRLF
				If iapiaction = "PRICE" Then
					strSQL = strSQL & " and apiAction in ('PRICE') " & VBCRLF
				ElseIf iapiaction = "EditSellYn" Then
					strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
				Else
					strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT') " & VBCRLF
				End If
				strSQL = strSQL & " and readdate is null " & VBCRLF
				strSQL = strSQL & " and lastUserid = 'system' "
				dbget.Execute strSQL
			End If
		End If
	End If
End Function

Function SugiQueLogInsert(imallid, iapiaction, iitemid, iresultcode, ilastErrMsg, ilastUpdateid)
	Dim strSQL

	ilastErrMsg = replace(ilastErrMsg, "'", "′")
	strSQL = ""
	strSQL = strSQL & " INSERT INTO [db_etcmall].[dbo].[tbl_outmall_API_Que] (mallid, apiAction, itemid, priority, regdate, readdate, findate, resultCode, lastErrMsg, lastUserid) VAlUES " & VBCRLF
	strSQL = strSQL & " ('"& imallid &"', '"& iapiaction &"', '"& iitemid &"', '999999', getdate(), getdate(), getdate(), '"& iresultcode &"', '"& LEFT(ilastErrMsg, 100) &"', '"& ilastUpdateid &"') " & VBCRLF
	dbget.Execute strSQL
	If iresultcode = "OK" Then
		If imallid = "interpark" OR imallid = "ezwel" OR imallid = "halfclub" Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
				If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "CHKSTAT") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
					If iapiaction = "EDIT" Then
						strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT', 'EDITBATCH') " & VBCRLF
                    ElseIf iapiaction = "CHKSTAT" Then
                        strSQL = strSQL & " and apiAction in ('CHKSTAT') " & VBCRLF
					Else
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		ElseIf imallid = "auction1010" Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" Then
				If (iapiaction = "EditInfo") OR (iapiaction = "EditSellYn") OR (iapiaction = "EDIT") OR (iapiaction = "PRICE") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
					If iapiaction = "EditInfo" Then
						strSQL = strSQL & " and apiAction in ('PRICE', 'EDITBATCH') " & VBCRLF
                    ElseIf iapiaction = "EDIT" Then
                        strSQL = strSQL & " and apiAction in ('EDIT') " & VBCRLF
					Else
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		ElseIf (imallid = "lotteCom") OR (imallid = "lotteon") OR (imallid = "lotteimall") OR (imallid = "cjmall") OR (imallid = "gsshop") OR (imallid = "nvstorefarm") OR (imallid = "nvstoremoonbangu") OR (imallid = "Mylittlewhoopee") OR (imallid = "nvstoregift") or imallid = ("gmarket1010") OR imallid = ("11st1010") OR imallid = ("ssg") OR imallid = ("coupang") OR imallid = ("hmall1010") OR imallid = ("WMP") OR imallid = ("wmpfashion") OR imallid = ("lfmall") OR imallid = ("shintvshopping") OR imallid = ("skstoa") OR imallid = ("kakaostore") Then
			If ilastUpdateid = "kjy8517" OR ilastUpdateid = "icommang" OR ilastUpdateid = "yhj0613" Then
				If (iapiaction = "EDIT") OR (iapiaction = "EditSellYn") OR (iapiaction = "PRICE") OR (iapiaction = "CHKSTAT") OR (UCase(iapiaction) = "EDITINFO") OR (UCase(iapiaction) = "EDITPOLICY") Then
					strSQL = ""
					strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
					strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
					strSQL = strSQL & " ,findate=getdate() " & VBCRLF
					strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
					strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
					strSQL = strSQL & " WHERE mallid = '"&imallid&"' " & VBCRLF
					strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
                    If (imallid = "gsshop") AND (iapiaction = "EDITINFO") Then
						strSQL = strSQL & " and apiAction in ('EDITINFO') " & VBCRLF
                    ElseIf (imallid = "gmarket1010") AND (UCase(iapiaction) = "EDITPOLICY") Then
                        strSQL = strSQL & " and apiAction in ('EDITPOLICY', 'EDITBATCH') " & VBCRLF
					ElseIf iapiaction = "EditSellYn" Then
						strSQL = strSQL & " and apiAction in ('SOLDOUT') " & VBCRLF
                    ElseIf iapiaction = "EDIT" Then
                        If imallid = "coupang" OR imallid = "ssg" OR imallid = "WMP" OR imallid = "wmpfashion" OR imallid = "11st1010" OR imallid = "lotteon" OR imallid = "gsshop" Then
                            strSQL = strSQL & " and apiAction in ('EDIT', 'EDITBATCH') " & VBCRLF
                        ElseIf imallid = "lfmall" Then
                            strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE') " & VBCRLF
                        Else
    						strSQL = strSQL & " and apiAction in ('EDIT') " & VBCRLF
                        End If
					ElseIf iapiaction = "CHKSTAT" Then
                        If imallid = "cjmall" Then
                            strSQL = strSQL & " and apiAction in ('CHKSTAT', 'CONFIRM') " & VBCRLF
                        Else
    						strSQL = strSQL & " and apiAction in ('CHKSTAT') " & VBCRLF
                        End If
					Else
						strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT') " & VBCRLF
					End If
					strSQL = strSQL & " and readdate is null " & VBCRLF
					strSQL = strSQL & " and lastUserid = 'system' "
					dbget.Execute strSQL
				End If
			End If
		End If
	End If
End Function

function DrawApiMallSelect(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >선택"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >롯데닷컴"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >롯데iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >인터파크"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='gseshop' "& chkIIF(selsitename="gseshop","selected","") &" >gseshop"
	buf = buf&"<option value='homeplus' "& chkIIF(selsitename="homeplus","selected","") &" >homeplus"
	buf = buf&"<option value='auction1010' "& chkIIF(selsitename="auction1010","selected","") &" >옥션"
	buf = buf&"</select>"

	response.write buf
end function

function DrawApiMallSelectSongjangInput(sitename,selsitename)
    dim buf
    buf = "<select name='"&sitename&"' >"
    buf = buf&"<option value=''  >선택"
	buf = buf&"<option value='lotteCom' "& chkIIF(selsitename="lotteCom","selected","") &" >롯데닷컴"
	buf = buf&"<option value='lotteimall' "& chkIIF(selsitename="lotteimall","selected","") &" >롯데iMall"
	buf = buf&"<option value='interpark' "& chkIIF(selsitename="interpark","selected","") &" >인터파크"
	buf = buf&"<option value='cjmall' "& chkIIF(selsitename="cjmall","selected","") &" >cjmall"
	buf = buf&"<option value='shoplinker' "& chkIIF(selsitename="shoplinker","selected","") &" >shoplinker"
	buf = buf&"</select>"

	response.write buf
end function

''어디에사용?
function DrawApiMallCheck()
    dim buf
    buf = ""
    buf = buf&"<input type='checkbox' name='outmallck' value='interpark'>인터파크"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteCom'>롯데닷컴"
    buf = buf&"<input type='checkbox' name='outmallck' value='lotteimall'>롯데iMall"

    response.write buf
end function

'' ssg어드민 업체출고관리  소스보기
function TenDlvCode2SSGDlvCode(itenCode)
''<option value="0000033023">SC로지스</option> ? 사가와익스프레스
''<option value="0000033026">경기택배</option>
''<option value="0000033029">네덱스택배</option>
''<option value="0000033032">대한통운</option> = >CJ대한통운
''<option value="0000033033">KG로지스(동부택배,옐로우캡)</option>
''<option value="0000033034">동원택배</option>
'<option value="0000033050">우체국EMS</option>
'<option value="0000033051">우체국등기</option>
'<option value="0000033052">우체국택배</option>
'<option value="0000033063">코덱스</option>
'<option value="0000033064">퀵/콜벤</option>
''<option value="0008369131">편의점택배</option>
    select Case itenCode
        CASE "1" : TenDlvCode2SSGDlvCode = "0000033071"     ''한진
        CASE "2" : TenDlvCode2SSGDlvCode = "0000033073"     ''현대 -> 롯데택배로 변경됨 2017-03-13 김진영 수정
        CASE "3" : TenDlvCode2SSGDlvCode = "0000033011"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2SSGDlvCode = "0000033011"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "5" : TenDlvCode2SSGDlvCode = "0000033056"     ''이클라인
        CASE "6" : TenDlvCode2SSGDlvCode = ""     ''삼성 HTH
        CASE "7" : TenDlvCode2SSGDlvCode = "0000033033"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2SSGDlvCode = "0000033052"     ''우체국택배
        CASE "9" : TenDlvCode2SSGDlvCode = "0000033017"     ''KGB택배
        CASE "10" : TenDlvCode2SSGDlvCode = "0000033035"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2SSGDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2SSGDlvCode = "0000033069"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2SSGDlvCode = "0000033033"     ''옐로우캡
        CASE "14" : TenDlvCode2SSGDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2SSGDlvCode = "0000033061"     ''중앙택배
        CASE "16" : TenDlvCode2SSGDlvCode = "0000033060"     ''주코택배
        CASE "17" : TenDlvCode2SSGDlvCode = "0000033067"     ''트라넷택배
        CASE "18" : TenDlvCode2SSGDlvCode = "0000033036"     ''로젠택배
        CASE "19" : TenDlvCode2SSGDlvCode = "0000033018"     ''KGB특급택배
        CASE "20" : TenDlvCode2SSGDlvCode = "0000033021"     ''KT로지스
        CASE "21" : TenDlvCode2SSGDlvCode = "0000033027"     ''경동택배
        CASE "22" : TenDlvCode2SSGDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2SSGDlvCode = ""     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2SSGDlvCode = ""     ''사가와익스프레스
        CASE "25" : TenDlvCode2SSGDlvCode = "0000033068"     ''하나로택배
        CASE "26" : TenDlvCode2SSGDlvCode = "0000033057"     ''일양택배
        CASE "27" : TenDlvCode2SSGDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2SSGDlvCode = "0000033033"     ''동부익스프레스
        CASE "29" : TenDlvCode2SSGDlvCode = "0000033025"     ''건영택배
        CASE "30" : TenDlvCode2SSGDlvCode = "0000033055"     ''이노지스
        CASE "31" : TenDlvCode2SSGDlvCode = "0000033062"     ''천일택배
        CASE "33" : TenDlvCode2SSGDlvCode = "0000033077"     ''호남택배
        CASE "34" : TenDlvCode2SSGDlvCode = "0000033030"     ''대신화물택배
        CASE "35" : TenDlvCode2SSGDlvCode = "0000033013"     ''CVSnet택배  -
        CASE "37" : TenDlvCode2SSGDlvCode = "0000038977"     ''합동택배  -
        CASE "38" : TenDlvCode2SSGDlvCode = "0000033014"     ''GTX로지스
        CASE "39" : TenDlvCode2SSGDlvCode = "0000033033"     ''KG로지스 - 동부익스프레스
        CASE "98" : TenDlvCode2SSGDlvCode = "0000033064"     ''퀵서비스->직배송
        CASE "41" : TenDlvCode2SSGDlvCode = "0000033033"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2SSGDlvCode = "0008369131"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "99" : TenDlvCode2SSGDlvCode = "0000033028"     ''기타  0000033028
        CASE  Else
            TenDlvCode2SSGDlvCode = "기타"      ''기타발송
    end Select
end function

Function TenDlvCode2WMPDlvCode(itenCode)
' D001 : 우체국택배
' D002 : CJ대한통운
' D003 : 한진택배
' D005 : 롯데택배
' D004 : 로젠택배
' D006 : KGB택배
' D011 : GTX로지스
' D007 : 일양로지스
' D008 : EMS
' D009 : DHL
' D010 : UPS
' D012 : 한의사랑택배
' D013 : 천일택배
' D014 : 건영택배
' D015 : 고려택배
' D016 : 한덱스
' D017 : Fedex
' D018 : 대신택배
' D019 : 경동택배
' D020 : CVSnet 편의점택배
' D021 : TNT Express
' D040 : CU 편의점택배
' D022 : USPS
' D041 : 농협택배
' D023 : TPL
' D042 : 세방
' D043 : 삼성전자 물류
' D024 : GSMNtoN
' D025 : 에어보이익스프레스
' D026 : KGL네크웍스
' D027 : 합동택배
' D028 : DHL Global Mail
' D029 : i-Parcel
' D030 : 포시즌 익스프레스
' D031 : 범한판토스
' D032 : APEX(ECMS Express)
' D034 : 굿투럭
' D035 : GSI Express
' D036 : CJ대한통운 국제특송
' D037 : SLX
' D038 : 호남택배
' D039 : 현대택배 해외특송
' D046 : GPS Logix
' D045 : 홈픽택배
' D044 : LG전자 물류
	SELECT Case itenCode
		CASE "1" : TenDlvCode2WMPDlvCode = "D003"		''한진
		CASE "2" : TenDlvCode2WMPDlvCode = "D005"		''현대 -> 롯데
		CASE "3" : TenDlvCode2WMPDlvCode = "D002"		''대한통운
		CASE "4" : TenDlvCode2WMPDlvCode = "D002"		''CJ GLS
		CASE "8" : TenDlvCode2WMPDlvCode = "D001"		''우체국택배
		CASE "9" : TenDlvCode2WMPDlvCode = "D006"		''KGB택배
		CASE "18" : TenDlvCode2WMPDlvCode = "D004"		''로젠택배
		CASE "21" : TenDlvCode2WMPDlvCode = "D019"		''경동택배
        CASE "22" : TenDlvCode2WMPDlvCode = "D015"		''고려택배
		CASE "26" : TenDlvCode2WMPDlvCode = "D007"		''일양택배
		CASE "29" : TenDlvCode2WMPDlvCode = "D014"		''건영택배
		CASE "31" : TenDlvCode2WMPDlvCode = "D013"		''천일택배
		CASE "33" : TenDlvCode2WMPDlvCode = "D038"		''호남택배
		CASE "34" : TenDlvCode2WMPDlvCode = "D018"		''대신화물택배
		CASE "35" : TenDlvCode2WMPDlvCode = "D020"      ''CVSnet택배
		CASE "37" : TenDlvCode2WMPDlvCode = "D027"		''합동택배  -
        CASE "38" : TenDlvCode2WMPDlvCode = "D011"		''GTX로지스
        CASE "42" : TenDlvCode2WMPDlvCode = "D040"		''CU Post => 대한통운
		CASE "91" : TenDlvCode2WMPDlvCode = "D009"      ''DHL
        CASE  Else
            TenDlvCode2WMPDlvCode = "0"     ''미입력
	End Select
end function

Function TenDlvCode2SabangNetDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2SabangNetDlvCode = "004"			''한진
		CASE "2" : TenDlvCode2SabangNetDlvCode = "002"			''현대 -> 롯데
		CASE "3" : TenDlvCode2SabangNetDlvCode = "001"			''대한통운
		CASE "4" : TenDlvCode2SabangNetDlvCode = "003"			''CJ GLS
		CASE "8" : TenDlvCode2SabangNetDlvCode = "009"			''우체국택배
		CASE "9" : TenDlvCode2SabangNetDlvCode = "005"		''KGB택배
		CASE "10" : TenDlvCode2SabangNetDlvCode = "021"		''아주택배 / 로엑스(구 아주)
		CASE "13" : TenDlvCode2SabangNetDlvCode = "032"		''옐로우캡
		CASE "17" : TenDlvCode2SabangNetDlvCode = "022"		''트라넷택배
		CASE "18" : TenDlvCode2SabangNetDlvCode = "007"		''로젠택배
		CASE "19" : TenDlvCode2SabangNetDlvCode = "035"		''KGB특급택배
		CASE "21" : TenDlvCode2SabangNetDlvCode = "013"		''경동택배
		CASE "22" : TenDlvCode2SabangNetDlvCode = "044"		''고려택배
		CASE "25" : TenDlvCode2SabangNetDlvCode = "010"		''하나로택배
		CASE "26" : TenDlvCode2SabangNetDlvCode = "047"		''일양택배
		CASE "27" : TenDlvCode2SabangNetDlvCode = "011"		''LOEX택배
		CASE "29" : TenDlvCode2SabangNetDlvCode = "043"		''건영택배
		CASE "30" : TenDlvCode2SabangNetDlvCode = "023"		''이노지스
		CASE "31" : TenDlvCode2SabangNetDlvCode = "016"		''천일택배
		CASE "33" : TenDlvCode2SabangNetDlvCode = "033"		''호남택배
		CASE "34" : TenDlvCode2SabangNetDlvCode = "037"		''대신화물택배
		CASE "37" : TenDlvCode2SabangNetDlvCode = "056"		''합동택배  -
		CASE "38" : TenDlvCode2SabangNetDlvCode = "053"		''GTX로지스
		CASE "41" : TenDlvCode2SabangNetDlvCode = "104"		''드림택배(동부택배,옐로우캡)  ''2018/02/13
		CASE "98" : TenDlvCode2SabangNetDlvCode = "999"		''퀵서비스->직배송
        CASE "99" : TenDlvCode2SabangNetDlvCode = "999"		''기타
		CASE  Else
			TenDlvCode2SabangNetDlvCode = "0"      ''미입력
	end Select
end function

function TenDlvCode2GmarketDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2GmarketDlvCode = "한진택배"     ''한진
        CASE "2" : TenDlvCode2GmarketDlvCode = "롯데택배"     ''현대 -> 롯데택배로 변경됨 2017-03-13 김진영 수정
        CASE "3" : TenDlvCode2GmarketDlvCode = "대한통운"     ''대한통운
        CASE "4" : TenDlvCode2GmarketDlvCode = "대한통운"     ''CJ GLS
        CASE "5" : TenDlvCode2GmarketDlvCode = "기타"     ''이클라인
        CASE "6" : TenDlvCode2GmarketDlvCode = "기타"     ''삼성 HTH
        CASE "7" : TenDlvCode2GmarketDlvCode = "기타"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2GmarketDlvCode = "우체국택배"     ''우체국택배
        CASE "9" : TenDlvCode2GmarketDlvCode = "KGB택배"     ''KGB택배
        CASE "10" : TenDlvCode2GmarketDlvCode = "기타"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2GmarketDlvCode = "기타"     ''오렌지택배
        CASE "12" : TenDlvCode2GmarketDlvCode = "한국택배"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2GmarketDlvCode = "옐로우캡택배"     ''옐로우캡
        CASE "14" : TenDlvCode2GmarketDlvCode = "기타"     ''나이스택배
        CASE "15" : TenDlvCode2GmarketDlvCode = "기타"     ''중앙택배
        CASE "16" : TenDlvCode2GmarketDlvCode = "기타"     ''주코택배
        CASE "17" : TenDlvCode2GmarketDlvCode = "기타"     ''트라넷택배
        CASE "18" : TenDlvCode2GmarketDlvCode = "로젠택배"     ''로젠택배
        CASE "19" : TenDlvCode2GmarketDlvCode = "KGB택배"     ''KGB특급택배
        CASE "20" : TenDlvCode2GmarketDlvCode = "기타"     ''KT로지스
        CASE "21" : TenDlvCode2GmarketDlvCode = "경동택배"     ''경동택배
        CASE "22" : TenDlvCode2GmarketDlvCode = "기타"     ''고려택배
        CASE "23" : TenDlvCode2GmarketDlvCode = "기타"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2GmarketDlvCode = "기타"     ''사가와익스프레스
        CASE "25" : TenDlvCode2GmarketDlvCode = "기타"     ''하나로택배
        CASE "26" : TenDlvCode2GmarketDlvCode = "일양택배"     ''일양택배
        CASE "27" : TenDlvCode2GmarketDlvCode = "기타"     ''LOEX택배
        CASE "28" : TenDlvCode2GmarketDlvCode = "기타"     ''동부익스프레스
        CASE "29" : TenDlvCode2GmarketDlvCode = "건영택배"     ''건영택배
        CASE "30" : TenDlvCode2GmarketDlvCode = "기타"     ''이노지스
        CASE "31" : TenDlvCode2GmarketDlvCode = "천일택배"     ''천일택배
        CASE "33" : TenDlvCode2GmarketDlvCode = "호남택배"     ''호남택배
        CASE "34" : TenDlvCode2GmarketDlvCode = "대신택배"     ''대신화물택배
        CASE "35" : TenDlvCode2GmarketDlvCode = "편의점택배(GS25)" ''"CVSNET(편의점)"     ''CVSnet택배  - 99 기타중소형택배
        CASE "38" : TenDlvCode2GmarketDlvCode = "GTX로지스"     ''GTX로지스
        CASE "42" : TenDlvCode2GmarketDlvCode = "편의점택배(GS25)"     ''CU편의점택배
        'CASE "39" : TenDlvCode2GmarketDlvCode = "KG로지스"     ''KG로지스 - 동부익스프레스
        CASE "39" : TenDlvCode2GmarketDlvCode = "드림택배"     ''2018-02-23 진영 수정
        CASE "98" : TenDlvCode2GmarketDlvCode = "퀵서비스"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2GmarketDlvCode = "직접배송"     ''기타
        CASE "102" : TenDlvCode2GmarketDlvCode = "직접배송"     ''직배송
        CASE  Else
            TenDlvCode2GmarketDlvCode = "기타"      ''기타발송
    end Select
end function

function TenDlvCode2AuctionDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2AuctionDlvCode = "hanjin"     ''한진
        CASE "2" : TenDlvCode2AuctionDlvCode = "hyundai"     ''현대
        CASE "3" : TenDlvCode2AuctionDlvCode = "korex"     ''대한통운
        CASE "4" : TenDlvCode2AuctionDlvCode = "cjgls"     ''CJ GLS
        CASE "5" : TenDlvCode2AuctionDlvCode = "etc"     ''이클라인
        CASE "6" : TenDlvCode2AuctionDlvCode = "samsung"     ''삼성 HTH
        CASE "7" : TenDlvCode2AuctionDlvCode = "dongbu"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2AuctionDlvCode = "epost"     ''우체국택배
        CASE "9" : TenDlvCode2AuctionDlvCode = "kgbls"     ''KGB택배
        CASE "10" : TenDlvCode2AuctionDlvCode = "etc"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2AuctionDlvCode = "etc"     ''오렌지택배
        CASE "12" : TenDlvCode2AuctionDlvCode = "etc"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2AuctionDlvCode = "yellow"     ''옐로우캡
        CASE "14" : TenDlvCode2AuctionDlvCode = "etc"     ''나이스택배
        CASE "15" : TenDlvCode2AuctionDlvCode = "etc"     ''중앙택배
        CASE "16" : TenDlvCode2AuctionDlvCode = "etc"     ''주코택배
        CASE "17" : TenDlvCode2AuctionDlvCode = "etc"     ''트라넷택배
        CASE "18" : TenDlvCode2AuctionDlvCode = "kgb"     ''로젠택배
        CASE "19" : TenDlvCode2AuctionDlvCode = "kgb"     ''KGB특급택배
        CASE "20" : TenDlvCode2AuctionDlvCode = "etc"     ''KT로지스
        CASE "21" : TenDlvCode2AuctionDlvCode = "kyungdong"     ''경동택배
        CASE "22" : TenDlvCode2AuctionDlvCode = "etc"     ''고려택배
        CASE "23" : TenDlvCode2AuctionDlvCode = "etc"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2AuctionDlvCode = "sagawa"     ''사가와익스프레스
        CASE "25" : TenDlvCode2AuctionDlvCode = "etc"     ''하나로택배
        CASE "26" : TenDlvCode2AuctionDlvCode = "ilyang"     ''일양택배
        CASE "27" : TenDlvCode2AuctionDlvCode = "etc"     ''LOEX택배
        CASE "28" : TenDlvCode2AuctionDlvCode = "dongbu"     ''동부익스프레스
        CASE "29" : TenDlvCode2AuctionDlvCode = "etc"     ''건영택배
        CASE "30" : TenDlvCode2AuctionDlvCode = "etc"     ''이노지스
        CASE "31" : TenDlvCode2AuctionDlvCode = "chonil"     ''천일택배
        CASE "33" : TenDlvCode2AuctionDlvCode = "etc"     ''호남택배
        CASE "34" : TenDlvCode2AuctionDlvCode = "daesin"     ''대신화물택배
        CASE "35" : TenDlvCode2AuctionDlvCode = "cvsnet"     ''CVSnet택배  - 99 기타중소형택배
        CASE "38" : TenDlvCode2AuctionDlvCode = "gtx"     ''GTX로지스
        CASE "39" : TenDlvCode2AuctionDlvCode = "dongbu"     ''KG로지스 - 동부익스프레스
        CASE "42" : TenDlvCode2AuctionDlvCode = "cvsnet"     ''CVSnet택배  - 99 기타중소형택배
        CASE "98" : TenDlvCode2AuctionDlvCode = "etc"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2AuctionDlvCode = "etc"     ''기타
        CASE  Else
            TenDlvCode2AuctionDlvCode = "etc"      ''기타발송
    end Select
end function

Function TenDlvCode2NvstorefarmDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2NvstorefarmDlvCode = "HANJIN"     ''한진
        CASE "2" : TenDlvCode2NvstorefarmDlvCode = "HYUNDAI"     ''현대
        CASE "3" : TenDlvCode2NvstorefarmDlvCode = "CJGLS"     ''대한통운
        CASE "4" : TenDlvCode2NvstorefarmDlvCode = "CJGLS"     ''CJ GLS
        CASE "5" : TenDlvCode2NvstorefarmDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2NvstorefarmDlvCode = ""     ''삼성 HTH
        CASE "7" : TenDlvCode2NvstorefarmDlvCode = "DONGBU"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2NvstorefarmDlvCode = "EPOST"     ''우체국택배
        CASE "9" : TenDlvCode2NvstorefarmDlvCode = "KGBLS"     ''KGB택배
        CASE "10" : TenDlvCode2NvstorefarmDlvCode = ""     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2NvstorefarmDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2NvstorefarmDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2NvstorefarmDlvCode = "YELLOW"     ''옐로우캡
        CASE "14" : TenDlvCode2NvstorefarmDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2NvstorefarmDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2NvstorefarmDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2NvstorefarmDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2NvstorefarmDlvCode = "KGB"     ''로젠택배
        CASE "19" : TenDlvCode2NvstorefarmDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2NvstorefarmDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2NvstorefarmDlvCode = "KDEXP"     ''경동택배
        CASE "22" : TenDlvCode2NvstorefarmDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2NvstorefarmDlvCode = ""     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2NvstorefarmDlvCode = ""     ''사가와익스프레스
        CASE "25" : TenDlvCode2NvstorefarmDlvCode = ""     ''하나로택배
        CASE "26" : TenDlvCode2NvstorefarmDlvCode = "ILYANG"     ''일양택배
        CASE "27" : TenDlvCode2NvstorefarmDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2NvstorefarmDlvCode = "DONGBU"     ''동부익스프레스
        CASE "29" : TenDlvCode2NvstorefarmDlvCode = "KUNYOUNG"     ''건영택배
        CASE "30" : TenDlvCode2NvstorefarmDlvCode = ""     ''이노지스
        CASE "31" : TenDlvCode2NvstorefarmDlvCode = "CHUNIL"     ''천일택배
        CASE "33" : TenDlvCode2NvstorefarmDlvCode = "HONAM"     ''호남택배
        CASE "34" : TenDlvCode2NvstorefarmDlvCode = "DAESIN"     ''대신화물택배
        CASE "35" : TenDlvCode2NvstorefarmDlvCode = "CVSNET"     ''CVSnet택배  - 99 기타중소형택배
        CASE "37" : TenDlvCode2NvstorefarmDlvCode = "HDEXP"     ''합동택배
        CASE "38" : TenDlvCode2NvstorefarmDlvCode = "INNOGIS"     ''GTX로지스   ''GTX(스카이로지스)::2586778  ''2015/06/29 추가
        CASE "42" : TenDlvCode2NvstorefarmDlvCode = "CUPARCEL"     ''CU편의점택배
        CASE "39" : TenDlvCode2NvstorefarmDlvCode = "DONGBU"     ''KG로지스
        CASE "98" : TenDlvCode2NvstorefarmDlvCode = "ETC1"     ''퀵서비스->직배송 | 2019-04-11 김진영..ETC1 추가 후 뒷단 처리
        CASE "99" : TenDlvCode2NvstorefarmDlvCode = "ETC2"     ''기타 | 2019-04-11 김진영..ETC2 추가 후 뒷단 처리
        CASE  Else
            TenDlvCode2NvstorefarmDlvCode = "CH1"      ''기타발송
    end Select
end function

function TenDlvCode2HomeplusDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2HomeplusDlvCode = "한진택배"     ''한진
        CASE "2" : TenDlvCode2HomeplusDlvCode = "현대택배"     ''현대
        CASE "3" : TenDlvCode2HomeplusDlvCode = "대한통운"     ''대한통운
        CASE "4" : TenDlvCode2HomeplusDlvCode = "CJGLS"     ''CJ GLS
        CASE "5" : TenDlvCode2HomeplusDlvCode = "이클라인택배"     ''이클라인
        CASE "6" : TenDlvCode2HomeplusDlvCode = "CJHTH"     ''삼성 HTH
        CASE "7" : TenDlvCode2HomeplusDlvCode = "훼미리택배"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2HomeplusDlvCode = "우체국택배"     ''우체국택배
        CASE "9" : TenDlvCode2HomeplusDlvCode = "KGB택배"     ''KGB택배
        CASE "10" : TenDlvCode2HomeplusDlvCode = "아주택배"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2HomeplusDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2HomeplusDlvCode = "한국택배"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2HomeplusDlvCode = "옐로우캡"     ''옐로우캡
        CASE "14" : TenDlvCode2HomeplusDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2HomeplusDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2HomeplusDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2HomeplusDlvCode = "트라넷택배"     ''트라넷택배
        CASE "18" : TenDlvCode2HomeplusDlvCode = "로젠택배"     ''로젠택배
        CASE "19" : TenDlvCode2HomeplusDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2HomeplusDlvCode = "KT로지스택배"     ''KT로지스
        CASE "21" : TenDlvCode2HomeplusDlvCode = "경동택배"     ''경동택배
        CASE "22" : TenDlvCode2HomeplusDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2HomeplusDlvCode = "쎄덱스"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2HomeplusDlvCode = "사가와택배"     ''사가와익스프레스
        CASE "25" : TenDlvCode2HomeplusDlvCode = "하나로택배"     ''하나로택배
        CASE "26" : TenDlvCode2HomeplusDlvCode = ""     ''일양택배
        CASE "27" : TenDlvCode2HomeplusDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2HomeplusDlvCode = "동부택배"     ''동부익스프레스
        CASE "29" : TenDlvCode2HomeplusDlvCode = ""     ''건영택배
        CASE "30" : TenDlvCode2HomeplusDlvCode = "이노지스택배"     ''이노지스
        CASE "31" : TenDlvCode2HomeplusDlvCode = "천일택배"     ''천일택배
        CASE "33" : TenDlvCode2HomeplusDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2HomeplusDlvCode = "대신택배"     ''대신화물택배
        CASE "35" : TenDlvCode2HomeplusDlvCode = "기타택배"     ''CVSnet택배  - 99 기타중소형택배
        CASE "98" : TenDlvCode2HomeplusDlvCode = "직배송"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2HomeplusDlvCode = "기타택배"     ''기타
        CASE  Else
            TenDlvCode2HomeplusDlvCode = "기타택배"      ''기타발송
    end Select
end function

function TenDlvCode2EzwelDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2EzwelDlvCode = "1016"     ''한진
        CASE "2" : TenDlvCode2EzwelDlvCode = "1017"     ''현대(롯데)
        CASE "3" : TenDlvCode2EzwelDlvCode = "1007"     ''대한통운
        CASE "4" : TenDlvCode2EzwelDlvCode = "1007"     ''CJ GLS
        CASE "8" : TenDlvCode2EzwelDlvCode = "1012"     ''우체국택배
        CASE "9" : TenDlvCode2EzwelDlvCode = "1002"     ''KGB택배
        CASE "13" : TenDlvCode2EzwelDlvCode = "1011"     ''옐로우캡
        CASE "18" : TenDlvCode2EzwelDlvCode = "1008"     ''로젠택배
        CASE "20" : TenDlvCode2EzwelDlvCode = "1082"     ''KT로지스
        CASE "21" : TenDlvCode2EzwelDlvCode = "1005"     ''경동택배
        CASE "24" : TenDlvCode2EzwelDlvCode = "1160"     ''사가와익스프레스
        CASE "26" : TenDlvCode2EzwelDlvCode = "1180"     ''일양택배
        CASE "28" : TenDlvCode2EzwelDlvCode = "1080"     ''동부익스프레스
        CASE "29" : TenDlvCode2EzwelDlvCode = "1106"     ''건영택배
        CASE "30" : TenDlvCode2EzwelDlvCode = "1163"     ''이노지스
        CASE "31" : TenDlvCode2EzwelDlvCode = "1014"     ''천일택배
        CASE "33" : TenDlvCode2EzwelDlvCode = "1107"     ''호남택배
        CASE "34" : TenDlvCode2EzwelDlvCode = "1200"     ''대신화물택배
        CASE "35" : TenDlvCode2EzwelDlvCode = "1007"     ''CVSnet택배  - 99 기타중소형택배 ''2019.08/20 송장입력불가 CJ로 변경(CJ로 추적됨)
        CASE "37" : TenDlvCode2EzwelDlvCode = "1102"     ''합동택배
        CASE "38" : TenDlvCode2EzwelDlvCode = "1260"     ''GTX로지스   ''GTX(스카이로지스)::2586778  ''2015/06/29 추가
        CASE "39" : TenDlvCode2EzwelDlvCode = "1080"     ''KG로지스
        CASE "41" : TenDlvCode2EzwelDlvCode = "1080"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "91" : TenDlvCode2EzwelDlvCode = "1001"
        CASE "98" : TenDlvCode2EzwelDlvCode = "1081"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2EzwelDlvCode = "1082"     ''기타
        CASE  Else
            TenDlvCode2EzwelDlvCode = "1082"      ''기타발송
    end Select
end function

function TenDlvCode2cjMallDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2cjMallDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2cjMallDlvCode = "11"     ''현대
        CASE "3" : TenDlvCode2cjMallDlvCode = "22"     ''대한통운
        CASE "4" : TenDlvCode2cjMallDlvCode = "22"     ''CJ GLS
        CASE "5" : TenDlvCode2cjMallDlvCode = "21"     ''이클라인
        CASE "6" : TenDlvCode2cjMallDlvCode = "29"     ''삼성 HTH
        CASE "7" : TenDlvCode2cjMallDlvCode = "79"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2cjMallDlvCode = "16"     ''우체국택배
        CASE "9" : TenDlvCode2cjMallDlvCode = "93"     ''KGB택배
        CASE "10" : TenDlvCode2cjMallDlvCode = "67"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2cjMallDlvCode = "17"     ''오렌지택배
        CASE "12" : TenDlvCode2cjMallDlvCode = "99"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2cjMallDlvCode = "69"     ''옐로우캡
        CASE "14" : TenDlvCode2cjMallDlvCode = "99"     ''나이스택배
        CASE "15" : TenDlvCode2cjMallDlvCode = "99"     ''중앙택배
        CASE "16" : TenDlvCode2cjMallDlvCode = "99"     ''주코택배
        CASE "17" : TenDlvCode2cjMallDlvCode = "57"     ''트라넷택배
        CASE "18" : TenDlvCode2cjMallDlvCode = "70"     ''로젠택배
        CASE "19" : TenDlvCode2cjMallDlvCode = "99"     ''KGB특급택배
        CASE "20" : TenDlvCode2cjMallDlvCode = "68"     ''KT로지스
        CASE "21" : TenDlvCode2cjMallDlvCode = "78"     ''경동택배
        CASE "22" : TenDlvCode2cjMallDlvCode = "99"     ''고려택배
        CASE "23" : TenDlvCode2cjMallDlvCode = "99"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2cjMallDlvCode = "62"     ''사가와익스프레스
        CASE "25" : TenDlvCode2cjMallDlvCode = "60"     ''하나로택배
        CASE "26" : TenDlvCode2cjMallDlvCode = "71"     ''일양택배
        CASE "27" : TenDlvCode2cjMallDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2cjMallDlvCode = "87"     ''동부익스프레스
        CASE "29" : TenDlvCode2cjMallDlvCode = "65"     ''건영택배
        CASE "30" : TenDlvCode2cjMallDlvCode = "88"     ''이노지스
        CASE "31" : TenDlvCode2cjMallDlvCode = "82"     ''천일택배
        CASE "33" : TenDlvCode2cjMallDlvCode = "58"     ''호남택배
        CASE "34" : TenDlvCode2cjMallDlvCode = "81"     ''대신화물택배
        CASE "35" : TenDlvCode2cjMallDlvCode = "12"     ''CVSnet택배  - CJ대한통운으로 정정요청
        CASE "39" : TenDlvCode2cjMallDlvCode = "87"     ''KG로지스
        CASE "98" : TenDlvCode2cjMallDlvCode = "32"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2cjMallDlvCode = "99"     ''기타
        CASE  Else
            TenDlvCode2cjMallDlvCode = "99"      ''기타발송
    end Select
end function

function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2InterParkDlvCode = "169178"     ''한진
        CASE "2" : TenDlvCode2InterParkDlvCode = "169198"     ''현대
        CASE "3" : TenDlvCode2InterParkDlvCode = "169177"     ''대한통운
        CASE "4" : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE "5" : TenDlvCode2InterParkDlvCode = "169211"     ''이클라인
        CASE "6" : TenDlvCode2InterParkDlvCode = "169181"     ''삼성 HTH
        CASE "7" : TenDlvCode2InterParkDlvCode = "231145"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2InterParkDlvCode = "169199"     ''우체국택배
        CASE "9" : TenDlvCode2InterParkDlvCode = "169187"     ''KGB택배
        CASE "10" : TenDlvCode2InterParkDlvCode = "169194"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2InterParkDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2InterParkDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2InterParkDlvCode = "169200"     ''옐로우캡
        CASE "14" : TenDlvCode2InterParkDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2InterParkDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2InterParkDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2InterParkDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2InterParkDlvCode = "169182"     ''로젠택배
        CASE "19" : TenDlvCode2InterParkDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2InterParkDlvCode = "169167"     ''KT로지스
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''경동택배
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''고려택배
        CASE "23" : TenDlvCode2InterParkDlvCode = "236288"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2InterParkDlvCode = "231491"     ''사가와익스프레스
        CASE "25" : TenDlvCode2InterParkDlvCode = "229381"     ''하나로택배
        CASE "26" : TenDlvCode2InterParkDlvCode = "263792"     ''일양택배
        CASE "27" : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX택배
        CASE "28" : TenDlvCode2InterParkDlvCode = "231145"     ''동부익스프레스
        CASE "29" : TenDlvCode2InterParkDlvCode = "231194"     ''건영택배
        CASE "30" : TenDlvCode2InterParkDlvCode = "266237"     ''이노지스
        CASE "31" : TenDlvCode2InterParkDlvCode = "230175"     ''천일택배
        CASE "33" : TenDlvCode2InterParkDlvCode = "250701"     ''호남택배
        CASE "34" : TenDlvCode2InterParkDlvCode = "258064"     ''대신화물택배
        CASE "35" : TenDlvCode2InterParkDlvCode = "169172"     ''CVSnet택배
        CASE "37" : TenDlvCode2InterParkDlvCode = "2641054"     ''합동택배
        CASE "38" : TenDlvCode2InterParkDlvCode = "2272970"     ''GTX로지스   ''GTX(스카이로지스)::2586778  ''2015/06/29 추가
        CASE "39" : TenDlvCode2InterParkDlvCode = "2964976"     ''KG로지스
        CASE "41" : TenDlvCode2InterParkDlvCode = "2964976"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
		CASE "42" : TenDlvCode2InterParkDlvCode = "169177"     ''CU Post => 대한통운
        CASE "50" : TenDlvCode2InterParkDlvCode = "4656462"     ''오늘의픽업
        CASE "54" : TenDlvCode2InterParkDlvCode = "169167"     ''기타
        CASE "98" : TenDlvCode2InterParkDlvCode = "169316"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2InterParkDlvCode = "169167"     ''기타
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""      ''기타발송(169167)
    end Select
end function

function TenDlvCode2LotteDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="99"

    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2LotteDlvCode = "27"     ''한진
        CASE "2" : TenDlvCode2LotteDlvCode = "1"     ''현대v
        CASE "3" : TenDlvCode2LotteDlvCode = "31"     ''대한통운
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''이클라인
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''우체국택배
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB택배
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteDlvCode = "37"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''중앙택배
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''로젠택배
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT로지스
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''경동택배
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''일양택배
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2LotteDlvCode = "35"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''건영택배
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''이노지스
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''천일택배
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''호남택배
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet택배
        CASE "39" : TenDlvCode2LotteDlvCode = "70"     ''KG로지스
        CASE "98" : TenDlvCode2LotteDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteDlvCode = "99"
    end Select
end function


'''롯데iMall 송장변환
function TenDlvCode2LotteiMallDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	이젠택배
''99	기타

    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2LotteiMallDlvCode = "11"     ''현대v
        CASE "3" : TenDlvCode2LotteiMallDlvCode = "12"     ''대한통운
        CASE "4" : TenDlvCode2LotteiMallDlvCode = "16"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2LotteiMallDlvCode = "22"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteiMallDlvCode = "31"     ''우체국택배
        CASE "9" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB택배
        CASE "10" : TenDlvCode2LotteiMallDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteiMallDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteiMallDlvCode = "37"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteiMallDlvCode = "32"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteiMallDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteiMallDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2LotteiMallDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteiMallDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteiMallDlvCode = "24"     ''로젠택배
        CASE "19" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteiMallDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''경동택배
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteiMallDlvCode = "47"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteiMallDlvCode = "43"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteiMallDlvCode = "46"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteiMallDlvCode = "18"     ''일양택배
        CASE "27" : TenDlvCode2LotteiMallDlvCode = "48"     ''LOEX택배
        CASE "28" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteiMallDlvCode = "99"     ''건영택배
        CASE "30" : TenDlvCode2LotteiMallDlvCode = "23"     ''이노지스
        CASE "31" : TenDlvCode2LotteiMallDlvCode = "17"     ''천일택배
        CASE "33" : TenDlvCode2LotteiMallDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2LotteiMallDlvCode = "38"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteiMallDlvCode = "99"     ''CVSnet택배
        CASE "98" : TenDlvCode2LotteiMallDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteiMallDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteiMallDlvCode = "99"
    end Select
end function

function TenDlvCode2LotteiMallNewDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	이젠택배
''99	기타
'rw itenCode
'response.end
    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallNewDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2LotteiMallNewDlvCode = "11"     ''현대v
        CASE "3" : TenDlvCode2LotteiMallNewDlvCode = "12"     ''대한통운
        CASE "4" : TenDlvCode2LotteiMallNewDlvCode = "12"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallNewDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteiMallNewDlvCode = "31"     ''우체국택배
        CASE "9" : TenDlvCode2LotteiMallNewDlvCode = "40"     ''KGB택배
        CASE "10" : TenDlvCode2LotteiMallNewDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteiMallNewDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteiMallNewDlvCode = "37"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteiMallNewDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteiMallNewDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2LotteiMallNewDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteiMallNewDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteiMallNewDlvCode = "24"     ''로젠택배
        CASE "19" : TenDlvCode2LotteiMallNewDlvCode = "40"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteiMallNewDlvCode = "37"     ''KT로지스
        CASE "21" : TenDlvCode2LotteiMallNewDlvCode = "49"     ''경동택배
        CASE "22" : TenDlvCode2LotteiMallNewDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteiMallNewDlvCode = "47"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteiMallNewDlvCode = "43"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteiMallNewDlvCode = "46"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteiMallNewDlvCode = "18"     ''일양택배
        CASE "27" : TenDlvCode2LotteiMallNewDlvCode = "48"     ''LOEX택배
        CASE "28" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''건영택배
        CASE "30" : TenDlvCode2LotteiMallNewDlvCode = "23"     ''이노지스
        CASE "31" : TenDlvCode2LotteiMallNewDlvCode = "17"     ''천일택배
        CASE "33" : TenDlvCode2LotteiMallNewDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2LotteiMallNewDlvCode = "38"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteiMallNewDlvCode = "62"     ''CVSnet택배
        CASE "39" : TenDlvCode2LotteiMallNewDlvCode = "21"     ''KG로지스
        CASE "98" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteiMallNewDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteiMallNewDlvCode = "99"
    end Select
end function

function TenDlvCode2GSShopDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="ZY"

    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2GSShopDlvCode = "HJ"     ''한진
        CASE "2" : TenDlvCode2GSShopDlvCode = "HD"     ''현대v
        CASE "3" : TenDlvCode2GSShopDlvCode = "DH"     ''대한통운
        CASE "4" : TenDlvCode2GSShopDlvCode = "DH"      ''"CJ"     ''CJ GLS  2017/07/27 CJ=>DH
        CASE "5" : TenDlvCode2GSShopDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2GSShopDlvCode = ""     ''삼성 HTH
        CASE "7" : TenDlvCode2GSShopDlvCode = "FA"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2GSShopDlvCode = "EP"     ''우체국택배
        CASE "9" : TenDlvCode2GSShopDlvCode = "KL"     ''KGB택배
        CASE "10" : TenDlvCode2GSShopDlvCode = ""     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2GSShopDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2GSShopDlvCode = ""     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2GSShopDlvCode = "YC"     ''옐로우캡
        CASE "14" : TenDlvCode2GSShopDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2GSShopDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2GSShopDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2GSShopDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2GSShopDlvCode = "KG"     ''로젠택배
        CASE "19" : TenDlvCode2GSShopDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2GSShopDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2GSShopDlvCode = "KD"     ''경동택배
        CASE "22" : TenDlvCode2GSShopDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2GSShopDlvCode = ""     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2GSShopDlvCode = ""     ''사가와익스프레스
        CASE "25" : TenDlvCode2GSShopDlvCode = ""     ''하나로택배v
        CASE "26" : TenDlvCode2GSShopDlvCode = "IY"     ''일양택배
        CASE "27" : TenDlvCode2GSShopDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2GSShopDlvCode = "FA"     ''동부익스프레스
        CASE "29" : TenDlvCode2GSShopDlvCode = ""     ''건영택배
        CASE "30" : TenDlvCode2GSShopDlvCode = ""     ''이노지스
        CASE "31" : TenDlvCode2GSShopDlvCode = "CI"     ''천일택배
        CASE "33" : TenDlvCode2GSShopDlvCode = "ZY"     ''호남택배
        CASE "34" : TenDlvCode2GSShopDlvCode = "DS"     ''대신화물택배
        CASE "35" : TenDlvCode2GSShopDlvCode = "CV"     ''CVSnet택배
        CASE "98" : TenDlvCode2GSShopDlvCode = "ZY"     ''퀵서비스
        CASE "99" : TenDlvCode2GSShopDlvCode = "ZY"     ''업체직송
        CASE  Else
            TenDlvCode2GSShopDlvCode = "ZY"
    end Select
end function

function LotteiMallDlvCode2Name(iltDlvCode)
    LotteiMallDlvCode2Name = "기타"
    if IsNULL(iltDlvCode) then Exit function
    iltDlvCode = TRIM(CStr(iltDlvCode))

    select Case iltDlvCode
        CASE "11" : LotteiMallDlvCode2Name="현대택배"
        CASE "12" : LotteiMallDlvCode2Name="씨제이대한통운"
        CASE "15" : LotteiMallDlvCode2Name="한진택배"
        CASE "16" : LotteiMallDlvCode2Name="CJGLS"
        CASE "17" : LotteiMallDlvCode2Name="천일택배"
        CASE "18" : LotteiMallDlvCode2Name="일양택배"
        CASE "19" : LotteiMallDlvCode2Name="기타택배"
        CASE "22" : LotteiMallDlvCode2Name="HTH택배"
        CASE "24" : LotteiMallDlvCode2Name="로젠택배"
        CASE "26" : LotteiMallDlvCode2Name="동부익스프레스"
        CASE "31" : LotteiMallDlvCode2Name="우체국택배"
        CASE "32" : LotteiMallDlvCode2Name="옐로우캡"
        CASE "34" : LotteiMallDlvCode2Name="아주택배"
        CASE "36" : LotteiMallDlvCode2Name="트라넷"
        CASE "37" : LotteiMallDlvCode2Name="한국택배"
        CASE "38" : LotteiMallDlvCode2Name="대신택배"
        CASE "40" : LotteiMallDlvCode2Name="KGB택배"
        CASE "41" : LotteiMallDlvCode2Name="이젠택배"
        CASE "43" : LotteiMallDlvCode2Name="사가와익스프레스"
        CASE "46" : LotteiMallDlvCode2Name="하나로택배"
        CASE "47" : LotteiMallDlvCode2Name="세덱스택배"
        CASE "48" : LotteiMallDlvCode2Name="로엑스택배"
        CASE "49" : LotteiMallDlvCode2Name="경동택배"
        CASE "99" : LotteiMallDlvCode2Name="기타"
        CASE  Else
            LotteiMallDlvCode2Name = "기타"
    end Select
end function

Function TenDlvCode211stDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode211stDlvCode = "00011"     ''한진
		CASE "2" : TenDlvCode211stDlvCode = "00012"    ''현대(롯데)택배
        CASE "3" : TenDlvCode211stDlvCode = "00034"     ''대한통운
        CASE "4" : TenDlvCode211stDlvCode = "00034"     ''CJ GLS
        CASE "8" : TenDlvCode211stDlvCode = "00007"     ''우체국택배
        CASE "18" : TenDlvCode211stDlvCode = "00002"     ''로젠택배
        CASE "21" : TenDlvCode211stDlvCode = "00026"     ''경동택배
        CASE "26" : TenDlvCode211stDlvCode = "00022"     ''일양택배
        CASE "29" : TenDlvCode211stDlvCode = "00037"     ''건영택배
        CASE "31" : TenDlvCode211stDlvCode = "00027"     ''천일택배
        CASE "37" : TenDlvCode211stDlvCode = "00035"     ''합동택배
        CASE "38" : TenDlvCode211stDlvCode = "00033"     ''GTX로지스   ''GTX(스카이로지스)::2586778  ''2015/06/29 추가
        CASE "39" : TenDlvCode211stDlvCode = "00001"     ''KG로지스 - 동부익스프레스
        CASE "99" : TenDlvCode211stDlvCode = "00099"     ''기타

        CASE "34" : TenDlvCode211stDlvCode = "00021"     ''대신(화물)택배
        CASE "35" : TenDlvCode211stDlvCode = "00060"     ''CVSnet택배
        CASE "42" : TenDlvCode211stDlvCode = "00061"     ''CU POST

        CASE  Else
            TenDlvCode211stDlvCode = "00099"      ''기타발송
    end Select
end function

Function TenDlvCode2HalfClubDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2HalfClubDlvCode = "3"			''한진
		CASE "2" : TenDlvCode2HalfClubDlvCode = "6"			''현대 -> 롯데
		CASE "3" : TenDlvCode2HalfClubDlvCode = "7"			''대한통운
		CASE "4" : TenDlvCode2HalfClubDlvCode = "1"			''CJ GLS
		CASE "8" : TenDlvCode2HalfClubDlvCode = "4"			''우체국택배
		CASE "9" : TenDlvCode2HalfClubDlvCode = "23"		''KGB택배
		CASE "10" : TenDlvCode2HalfClubDlvCode = "27"		''아주택배 / 로엑스(구 아주)
		CASE "13" : TenDlvCode2HalfClubDlvCode = "10"		''옐로우캡
		CASE "17" : TenDlvCode2HalfClubDlvCode = "16"		''트라넷택배
		CASE "18" : TenDlvCode2HalfClubDlvCode = "8"		''로젠택배
		CASE "19" : TenDlvCode2HalfClubDlvCode = "23"		''KGB특급택배
		CASE "21" : TenDlvCode2HalfClubDlvCode = "15"		''경동택배
		CASE "22" : TenDlvCode2HalfClubDlvCode = "24"		''고려택배
		CASE "25" : TenDlvCode2HalfClubDlvCode = "25"		''하나로택배
		CASE "26" : TenDlvCode2HalfClubDlvCode = "32"		''일양택배
		CASE "27" : TenDlvCode2HalfClubDlvCode = "27"		''LOEX택배
		CASE "29" : TenDlvCode2HalfClubDlvCode = "56"		''건영택배
		CASE "30" : TenDlvCode2HalfClubDlvCode = "30"		''이노지스
		CASE "31" : TenDlvCode2HalfClubDlvCode = "40"		''천일택배
		CASE "33" : TenDlvCode2HalfClubDlvCode = "54"		''호남택배
		CASE "34" : TenDlvCode2HalfClubDlvCode = "33"		''대신화물택배
		CASE "37" : TenDlvCode2HalfClubDlvCode = "37"		''합동택배  -
		CASE "38" : TenDlvCode2HalfClubDlvCode = "48"		''GTX로지스
		CASE "41" : TenDlvCode2HalfClubDlvCode = "26"		''드림택배(동부택배,옐로우캡)  ''2018/02/13
		CASE "98" : TenDlvCode2HalfClubDlvCode = "39"		''퀵서비스->직배송
		CASE  Else
			TenDlvCode2HalfClubDlvCode = "0"      ''미입력
	end Select
end function

Function TenDlvCode2CoupangDlvCode(itenCode)
	select Case itenCode
		CASE "1" : TenDlvCode2CoupangDlvCode = "HANJIN"			''한진
		CASE "2" : TenDlvCode2CoupangDlvCode = "HYUNDAI"			''현대 -> 롯데
		CASE "3" : TenDlvCode2CoupangDlvCode = "CJGLS"			''대한통운
		CASE "4" : TenDlvCode2CoupangDlvCode = "CJGLS"			''CJ GLS
		CASE "5" : TenDlvCode2CoupangDlvCode = "CSLOGIS"
		CASE "8" : TenDlvCode2CoupangDlvCode = "EPOST"			''우체국택배
		CASE "9" : TenDlvCode2CoupangDlvCode = "KGBPS"		''KGB택배
		CASE "10" : TenDlvCode2CoupangDlvCode = "AJOU"		''아주택배 / 로엑스(구 아주)
		CASE "18" : TenDlvCode2CoupangDlvCode = "KGB"		''로젠택배
		CASE "21" : TenDlvCode2CoupangDlvCode = "KDEXP"		''경동택배
		CASE "24" : TenDlvCode2CoupangDlvCode = "CSLOGIS"
		CASE "26" : TenDlvCode2CoupangDlvCode = "ILYANG"		''일양택배
		CASE "29" : TenDlvCode2CoupangDlvCode = "KUNYOUNG"		''건영택배
		CASE "31" : TenDlvCode2CoupangDlvCode = "CHUNIL"		''천일택배
		CASE "33" : TenDlvCode2CoupangDlvCode = "HONAM"		''호남택배
		CASE "34" : TenDlvCode2CoupangDlvCode = "DAESIN"		''대신화물택배
		CASE "35" : TenDlvCode2CoupangDlvCode = "CVS"
		CASE "36" : TenDlvCode2CoupangDlvCode = "HANJIN"
		CASE "37" : TenDlvCode2CoupangDlvCode = "HDEXP"		''합동택배  -
		CASE "39" : TenDlvCode2CoupangDlvCode = "DONGBU"
		CASE "41" : TenDlvCode2CoupangDlvCode = "DONGBU"
        CASE "42" : TenDlvCode2CoupangDlvCode = "BGF"       'CU POST를 BGF포스트로 변경
        CASE "47" : TenDlvCode2CoupangDlvCode = "TPMLOGIS"	''티피엠로지스(용달이특송)
        CASE "54" : TenDlvCode2CoupangDlvCode = "DIRECT"	'NDEX KOREA를 업체직송으로 요청..20220721 김은주B 요청
		CASE "91" : TenDlvCode2CoupangDlvCode = "DHL"
        CASE "98" : TenDlvCode2CoupangDlvCode = "DIRECT"    '퀵서비스를 업체직송으로 요청..20190411 하소라님 요청
		CASE "99" : TenDlvCode2CoupangDlvCode = "DIRECT"
	end Select
end function

Function TenDlvCode2HmallDlvCode(itenCode)
'1	11	롯데택배
'2	12	CJ대한통운
'3	13	한진택배
'4	16	자가배송
'5	24	굿스포스트
'6	25	현대리바트
'7	29	KGB택배
'8	31	우편등기
'9	33	로젠택배
'10	35	우체국택배
'11	38	일양택배
'12	60	퀵서비스
'13	61	경동택배
'14	63	한서호남택배
'15	64	천일택배
'16	65	편의점택배(GS25)
'17	68	합동택배
'18	69	대신택배
'19	70	건영택배
'20	71	GTX로지스
'21	72	한의사랑택배
'22	78	세방택배
'23	79	농협택배
'24	83	TNT express
'25	84	범한판토스
'26	89	DHL

	SELECT Case itenCode
		CASE "1" : TenDlvCode2HmallDlvCode = "13"		''한진
		CASE "2" : TenDlvCode2HmallDlvCode = "11"		''현대 -> 롯데
		CASE "3" : TenDlvCode2HmallDlvCode = "12"		''대한통운
		CASE "4" : TenDlvCode2HmallDlvCode = "12"		''CJ GLS
		CASE "8" : TenDlvCode2HmallDlvCode = "35"		''우체국택배
		CASE "9" : TenDlvCode2HmallDlvCode = "29"		''KGB택배
		CASE "18" : TenDlvCode2HmallDlvCode = "33"		''로젠택배
		CASE "21" : TenDlvCode2HmallDlvCode = "61"		''경동택배
		CASE "26" : TenDlvCode2HmallDlvCode = "38"		''일양택배
		CASE "29" : TenDlvCode2HmallDlvCode = "70"		''건영택배
		CASE "31" : TenDlvCode2HmallDlvCode = "64"		''천일택배
		CASE "33" : TenDlvCode2HmallDlvCode = "63"		''호남택배
		CASE "34" : TenDlvCode2HmallDlvCode = "69"		''대신화물택배
		CASE "35" : TenDlvCode2HmallDlvCode = "65"
		CASE "37" : TenDlvCode2HmallDlvCode = "68"		''합동택배  -
        CASE "38" : TenDlvCode2HmallDlvCode = "71"		''GTX로지스
        CASE "42" : TenDlvCode2HmallDlvCode = "65"		''CU POST
        CASE "47" : TenDlvCode2HmallDlvCode = "93"		''티피엠로지스
        CASE "45" : TenDlvCode2HmallDlvCode = "74"		''FLF퍼레버택배
		CASE "91" : TenDlvCode2HmallDlvCode = "89"
        CASE "98" : TenDlvCode2HmallDlvCode = "60"		''퀵서비스
        CASE "99" : TenDlvCode2HmallDlvCode = "16"
	End Select
End Function

Function TenDlvCode2Wetoo1300kDlvCode(itenCode)
' D001	CJ대한통운
' D003	퀵배송(오토바이)
' D004	KGB택배
' D005	일양택배
' D008	옐로우캡택배
' D009	기타
' D010	대한통운
' D011	로젠택배
' D015	등기우편
' D016	우체국택배
' D020	한진택배
' D021	롯데택배
' D023	동부익스프레스택배
' D025	이노지스택배
' D026	자체배송
' D027	천일택배
' D029	대신택배
' D030	경동택배
' D031	티켓(현장발권)
' D032	EMS(해외배송)
' D033	온라인 다운로드
' D034	건영택배
' D035	쿠폰발급
' D036	CVSNET(편의점)
' D037	합동택배
' D038	한우리물류
' D039	GTX로지스
' D040	드림택배
    select Case itenCode
        CASE "1" : TenDlvCode2Wetoo1300kDlvCode = "D020"     ''한진
        CASE "2" : TenDlvCode2Wetoo1300kDlvCode = "D021"     '롯데택배
        CASE "3" : TenDlvCode2Wetoo1300kDlvCode = "D001"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2Wetoo1300kDlvCode = "D001"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "8" : TenDlvCode2Wetoo1300kDlvCode = "D016"     ''우체국택배
        CASE "9" : TenDlvCode2Wetoo1300kDlvCode = "D004"     ''KGB택배
        CASE "13" : TenDlvCode2Wetoo1300kDlvCode = "D008"     ''옐로우캡
        CASE "18" : TenDlvCode2Wetoo1300kDlvCode = "D011"     ''로젠택배
        CASE "21" : TenDlvCode2Wetoo1300kDlvCode = "D030"     ''경동택배
        CASE "26" : TenDlvCode2Wetoo1300kDlvCode = "D005"     ''일양택배
        CASE "28" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''동부익스프레스
        CASE "29" : TenDlvCode2Wetoo1300kDlvCode = "D034"     ''건영택배
        CASE "30" : TenDlvCode2Wetoo1300kDlvCode = "D025"     ''이노지스
        CASE "31" : TenDlvCode2Wetoo1300kDlvCode = "D027"     ''천일택배
        CASE "34" : TenDlvCode2Wetoo1300kDlvCode = "D029"     ''대신화물택배
        CASE "35" : TenDlvCode2Wetoo1300kDlvCode = "D036"     ''CVSnet택배  -
        CASE "37" : TenDlvCode2Wetoo1300kDlvCode = "D037"     ''합동택배  -
        CASE "38" : TenDlvCode2Wetoo1300kDlvCode = "D039"     ''GTX로지스
        CASE "39" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''KG로지스 - 동부익스프레스
        CASE "40" : TenDlvCode2Wetoo1300kDlvCode = "D023"     ''KG로지스 - 동부익스프레스
        CASE "41" : TenDlvCode2Wetoo1300kDlvCode = "D040"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2Wetoo1300kDlvCode = "D036"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "90" : TenDlvCode2Wetoo1300kDlvCode = "D032"    'EMS	http://se
        CASE "98" : TenDlvCode2Wetoo1300kDlvCode = "D003"    '퀵서비스		Y
        CASE "99" : TenDlvCode2Wetoo1300kDlvCode = "D009"    '기타		Y	N
        CASE "100": TenDlvCode2Wetoo1300kDlvCode = "D038"     '한우리물류	h
        CASE  Else
            TenDlvCode2Wetoo1300kDlvCode = "D009"      ''기타발송
    end Select
End Function

Function TenDlvCode2LotteonDlvCode(itenCode)
' 0001||롯데택배
' 0002||CJ대한통운
' 0003||현대택배
' 0004||우체국택배
' 0005||로젠택배
' 0006||한진택배
' 0007||APEX(ECMS Express)
' 0008||DHL
' 0009||DHL Global Mail
' 0010||EMS
' 0011||Fedex
' 0012||GSI Express
' 0013||GSMNtoN(인로스)
' 0014||GTX 로지스 택배
' 0015||i-Parcel
' 0016||KGB택배
' 0017||KGL네트웍스
' 0018||KG로지스
' 0019||SEDEX
' 0020||TNT Express
' 0021||TPL
' 0022||USPS
' 0023||건영택배
' 0024||경동택배
' 0025||고려택배
' 0026||굿투럭
' 0027||대신정기화물택배
' 0028||대신택배
' 0029||대한통운
' 0030||동부익스프레스
' 0031||드림택배
' 0032||무지개오토특송
' 0033||범한판토스
' 0034||삼성HTH
' 0035||애니트랙
' 0036||에어보이익스프레스
' 0037||엘로우캡택배
' 0038||우체국
' 0039||우체국등기
' 0040||이노지스택배
' 0041||일양로지스
' 0042||일양택배
' 0043||천일택배
' 0044||편의점택배
' 0045||포시즌익스프레스
' 0046||하나로택배
' 0047||한덱스
' 0048||한의사랑택배
' 0049||합동택배
' 0050||호남택배
' 0051||KT로지스
' 0052||사가와
' 0053||우리택배
' 0054||제니엘
' 9000||자체배송
' 9999||기타택배
' LE_QUICK||엘롯데퀵배송사

    select Case itenCode
        CASE "1" : TenDlvCode2LotteonDlvCode = "0006"     ''한진
        CASE "2" : TenDlvCode2LotteonDlvCode = "0001"     ''현대 -> 롯데택배로 변경됨 2017-03-13 김진영 수정
        CASE "3" : TenDlvCode2LotteonDlvCode = "0002"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2LotteonDlvCode = "0002"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "5" : TenDlvCode2LotteonDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2LotteonDlvCode = "0034"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteonDlvCode = "0030"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2LotteonDlvCode = "0004"     ''우체국택배
        CASE "9" : TenDlvCode2LotteonDlvCode = "0016"     ''KGB택배
        CASE "10" : TenDlvCode2LotteonDlvCode = ""     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteonDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteonDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2LotteonDlvCode = "0016"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteonDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteonDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2LotteonDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteonDlvCode = "0051"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteonDlvCode = "0005"     ''로젠택배
        CASE "19" : TenDlvCode2LotteonDlvCode = "0016"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteonDlvCode = "0051"     ''KT로지스
        CASE "21" : TenDlvCode2LotteonDlvCode = "0024"     ''경동택배
        CASE "22" : TenDlvCode2LotteonDlvCode = "0025"     ''고려택배
        CASE "23" : TenDlvCode2LotteonDlvCode = ""     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteonDlvCode = "0052"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteonDlvCode = "0046"     ''하나로택배
        CASE "26" : TenDlvCode2LotteonDlvCode = "0041"     ''일양택배
        CASE "27" : TenDlvCode2LotteonDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2LotteonDlvCode = "0030"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteonDlvCode = "0023"     ''건영택배
        CASE "30" : TenDlvCode2LotteonDlvCode = "0040"     ''이노지스
        CASE "31" : TenDlvCode2LotteonDlvCode = "0043"     ''천일택배
        CASE "33" : TenDlvCode2LotteonDlvCode = "0050"     ''호남택배
        CASE "34" : TenDlvCode2LotteonDlvCode = "0028"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteonDlvCode = "0044"     ''CVSnet택배  -
        CASE "37" : TenDlvCode2LotteonDlvCode = "0049"     ''합동택배  -
        CASE "38" : TenDlvCode2LotteonDlvCode = "0014"     ''GTX로지스
        CASE "39" : TenDlvCode2LotteonDlvCode = "0018"     ''KG로지스 - 동부익스프레스
        CASE "98" : TenDlvCode2LotteonDlvCode = ""     ''퀵서비스->직배송
        CASE "41" : TenDlvCode2LotteonDlvCode = "0031"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2LotteonDlvCode = "0044"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "99" : TenDlvCode2LotteonDlvCode = "9999"     ''기타  0000033028
        CASE  Else
            TenDlvCode2LotteonDlvCode = "9999"      ''기타발송
    end Select
end function

Function TenDlvCode2ShintvshoppingDlvCode(itenCode)
' 10||CJ 대한통운
' 11||롯데택배
' 12||로젠택배
' 13||우체국택배
' 14||한진택배
' 17||경동택배
' 20||이노지스
' 21||일양택배
' 22||천일택배
' 23||로엑스(아주)택배
' 24||SC로지스택배
' 25||대신택배
' 26||CVS편의점택배
' 27||자체배송
' 28||설치상품
' 29||삼성HTH
' 30||훼미리넷
' 31||이클라인
' 32||주코택배
' 33||호남택배
' 34||우리택배
' 35||트라넷
' 36||한국택배
' 37||합동택배
' 38||GTX로지스
' 39||SLX택배
' 40||자체처리
' 41||로젝스
' 42||HI택배
' 43||퍼레버택배
' 44||YDH
' 45||화물을부탁해
' 60||대운글로벌
' 61||ACI Express
' 62||이투마스
' 63||에이스물류
' 64||캐나다쉬핑
' 65||지디에이코리아
' 66||바바바로지스
' 67||팀프레시
' 70||동원새벽배송
' 71||올타코리아
' 72||롯데칠성
' 73||yunda express
' 74||국제익스플레스
' 75||윈핸드해운상공
' 76||히얼위고
' 77||티피엠코리아
' 78||건영택배
' 79||핑퐁
' 90||미정
' 99||기타

    select Case itenCode
        CASE "1" : TenDlvCode2ShintvshoppingDlvCode = "14"     ''한진
        CASE "2" : TenDlvCode2ShintvshoppingDlvCode = "11"     '롯데택배
        CASE "3" : TenDlvCode2ShintvshoppingDlvCode = "10"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2ShintvshoppingDlvCode = "10"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "5" : TenDlvCode2ShintvshoppingDlvCode = "31"     ''이클라인
        CASE "5" : TenDlvCode2ShintvshoppingDlvCode = "24"     ''SC로지스
        CASE "8" : TenDlvCode2ShintvshoppingDlvCode = "13"     ''우체국택배
        CASE "9" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KGB택배
        CASE "10" : TenDlvCode2ShintvshoppingDlvCode = "23"     ''아주택배 / 로엑스(구 아주)
        CASE "12" : TenDlvCode2ShintvshoppingDlvCode = "36"     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2ShintvshoppingDlvCode = ""     ''옐로우캡
        CASE "16" : TenDlvCode2ShintvshoppingDlvCode = "32"     ''주코택배
        CASE "17" : TenDlvCode2ShintvshoppingDlvCode = "35"     ''트라넷택배
        CASE "18" : TenDlvCode2ShintvshoppingDlvCode = "12"     ''로젠택배
        CASE "20" : TenDlvCode2ShintvshoppingDlvCode = "99"     ''KT로지스..2023-02-02 김진영..매칭할 게 없음..기타로 일단 처리
        CASE "21" : TenDlvCode2ShintvshoppingDlvCode = "17"     ''경동택배
        CASE "22" : TenDlvCode2ShintvshoppingDlvCode = ""     ''고려택배
        CASE "24" : TenDlvCode2ShintvshoppingDlvCode = "24"     ''SC로지스
        CASE "25" : TenDlvCode2ShintvshoppingDlvCode = ""     ''하나로택배
        CASE "26" : TenDlvCode2ShintvshoppingDlvCode = "21"     ''일양택배
        CASE "27" : TenDlvCode2ShintvshoppingDlvCode = "23"     ''LOEX택배
        CASE "28" : TenDlvCode2ShintvshoppingDlvCode = ""     ''동부익스프레스
        CASE "29" : TenDlvCode2ShintvshoppingDlvCode = "78"     ''건영택배
        CASE "30" : TenDlvCode2ShintvshoppingDlvCode = "20"     ''이노지스
        CASE "31" : TenDlvCode2ShintvshoppingDlvCode = "22"     ''천일택배
        CASE "33" : TenDlvCode2ShintvshoppingDlvCode = "33"     ''호남택배
        CASE "34" : TenDlvCode2ShintvshoppingDlvCode = "25"     ''대신화물택배
        CASE "35" : TenDlvCode2ShintvshoppingDlvCode = "26"     ''CVSnet택배  -
        CASE "36" : TenDlvCode2ShintvshoppingDlvCode = ""     '한진정기화물
        CASE "37" : TenDlvCode2ShintvshoppingDlvCode = "37"     ''합동택배  -
        CASE "38" : TenDlvCode2ShintvshoppingDlvCode = "38"     ''GTX로지스
        CASE "39" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KG로지스 - 동부익스프레스
        CASE "40" : TenDlvCode2ShintvshoppingDlvCode = ""     ''KG로지스 - 동부익스프레스
        CASE "41" : TenDlvCode2ShintvshoppingDlvCode = ""     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2ShintvshoppingDlvCode = "26"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "43" : TenDlvCode2ShintvshoppingDlvCode = "42"    'HI택배	http:
        CASE "44" : TenDlvCode2ShintvshoppingDlvCode = ""    '홈픽	http://ww
        CASE "45" : TenDlvCode2ShintvshoppingDlvCode = "43"    'FLF퍼레버택배	h
        CASE "46" : TenDlvCode2ShintvshoppingDlvCode = ""    'FedEx	https
        CASE "47" : TenDlvCode2ShintvshoppingDlvCode = "77"    '티피엠코리아	h
        CASE "48" : TenDlvCode2ShintvshoppingDlvCode = ""    '로지스밸리	h
        CASE "49" : TenDlvCode2ShintvshoppingDlvCode = ""    '로지스밸리택배	h
        CASE "90" : TenDlvCode2ShintvshoppingDlvCode = ""    'EMS	http://se
        CASE "91" : TenDlvCode2ShintvshoppingDlvCode = ""    'DHL	http://ww
        CASE "98" : TenDlvCode2ShintvshoppingDlvCode = "99"    '퀵서비스		Y
        CASE "99" : TenDlvCode2ShintvshoppingDlvCode = "99"    '기타		Y	N
        CASE "100": TenDlvCode2ShintvshoppingDlvCode = ""     '한우리물류	h
        CASE  Else
            TenDlvCode2ShintvshoppingDlvCode = "99"      ''기타발송
    end Select
end function

Function TenDlvCode2SkstoaDlvCode(itenCode)
' 10||CJ대한통운
' 11||한진택배
' 12||롯데택배
' 13||우체국택배
' 14||로젠택배
' 18||일양로지스
' 26||SBGLS
' 27||대신택배
' 28||경동택배
' 29||합동택배
' 30||편의점택배
' 32||한의사랑택배
' 34||천일택배
' 35||건영택배
' 36||고려택배
' 37||티피엠코리아
' 38||시알로지텍
' 39||로지스밸리
' 40||자체처리
' 47||자체배송
' 48||설치상품
' 90||동원새벽택배
' 91||롯데글로벌
' 92||팀프레시
' 99||기타
    select Case itenCode
        CASE "1" : TenDlvCode2SkstoaDlvCode = "11"     ''한진
        CASE "2" : TenDlvCode2SkstoaDlvCode = "12"     '롯데택배
        CASE "3" : TenDlvCode2SkstoaDlvCode = "10"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2SkstoaDlvCode = "10"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "5" : TenDlvCode2SkstoaDlvCode = ""     ''이클라인
        CASE "5" : TenDlvCode2SkstoaDlvCode = ""     ''SC로지스
        CASE "8" : TenDlvCode2SkstoaDlvCode = "13"     ''우체국택배
        CASE "9" : TenDlvCode2SkstoaDlvCode = ""   ''KGB택배
        CASE "10" : TenDlvCode2SkstoaDlvCode = ""     ''아주택배 / 로엑스(구 아주)
        CASE "12" : TenDlvCode2SkstoaDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2SkstoaDlvCode = ""   ''옐로우캡
        CASE "16" : TenDlvCode2SkstoaDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2SkstoaDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2SkstoaDlvCode = "14"     ''로젠택배
        CASE "20" : TenDlvCode2SkstoaDlvCode = ""   ''KT로지스
        CASE "21" : TenDlvCode2SkstoaDlvCode = "28"     ''경동택배
        CASE "22" : TenDlvCode2SkstoaDlvCode = "36"   ''고려택배
        CASE "24" : TenDlvCode2SkstoaDlvCode = ""     ''SC로지스
        CASE "25" : TenDlvCode2SkstoaDlvCode = ""   ''하나로택배
        CASE "26" : TenDlvCode2SkstoaDlvCode = "18"     ''일양택배
        CASE "27" : TenDlvCode2SkstoaDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2SkstoaDlvCode = ""   ''동부익스프레스
        CASE "29" : TenDlvCode2SkstoaDlvCode = "35"     ''건영택배
        CASE "30" : TenDlvCode2SkstoaDlvCode = ""     ''이노지스
        CASE "31" : TenDlvCode2SkstoaDlvCode = "34"     ''천일택배
        CASE "33" : TenDlvCode2SkstoaDlvCode = "40"     ''호남택배 / 하소라님이 자체배송으로 요청
        CASE "34" : TenDlvCode2SkstoaDlvCode = "27"     ''대신화물택배
        CASE "35" : TenDlvCode2SkstoaDlvCode = "30"     ''CVSnet택배  -
        CASE "36" : TenDlvCode2SkstoaDlvCode = ""   '한진정기화물
        CASE "37" : TenDlvCode2SkstoaDlvCode = "29"     ''합동택배  -
        CASE "38" : TenDlvCode2SkstoaDlvCode = ""     ''GTX로지스
        CASE "39" : TenDlvCode2SkstoaDlvCode = ""   ''KG로지스 - 동부익스프레스
        CASE "40" : TenDlvCode2SkstoaDlvCode = ""   ''KG로지스 - 동부익스프레스
        CASE "41" : TenDlvCode2SkstoaDlvCode = ""   ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2SkstoaDlvCode = "30"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "43" : TenDlvCode2SkstoaDlvCode = ""    'HI택배	http:
        CASE "44" : TenDlvCode2SkstoaDlvCode = ""  '홈픽	http://ww
        CASE "45" : TenDlvCode2SkstoaDlvCode = ""    'FLF퍼레버택배	h
        CASE "46" : TenDlvCode2SkstoaDlvCode = ""  'FedEx	https
        CASE "47" : TenDlvCode2SkstoaDlvCode = "37"    '티피엠코리아	h
        CASE "48" : TenDlvCode2SkstoaDlvCode = "39"  '로지스밸리	h
        CASE "49" : TenDlvCode2SkstoaDlvCode = "39"  '로지스밸리택배	h
        CASE "90" : TenDlvCode2SkstoaDlvCode = ""  'EMS	http://se
        CASE "91" : TenDlvCode2SkstoaDlvCode = ""  'DHL	http://ww
        CASE "98" : TenDlvCode2SkstoaDlvCode = "99"    '퀵서비스		Y
        CASE "99" : TenDlvCode2SkstoaDlvCode = "99"    '기타		Y	N
        CASE "100": TenDlvCode2SkstoaDlvCode = ""   '한우리물류	h
        CASE  Else
            TenDlvCode2SkstoaDlvCode = "99"      ''기타발송
    end Select
end function

Function TenDlvCode2MarketforDlvCode(itenCode)
' korex       CJ 대한통운
' yellow      옐로우캡
' logen       로젠택배
' dongbu      동부익스프레스택배
' epost       우체국택배
' hanjin      한진택배
' hyundai     롯데택배(구 현대택배)
' kdexp       경동택배
' ETC         기타
' pantos      범한판토스
' hilogis     HI 택배
' tnt         TNT
' kgbps       KGB 택배
' chunil      천일택배
' ilyang      일양로지스
' fedex       FEDEX
' swgexp      성원글로벌
' daesin      대신택배
' ups         UPS
' hdexp       합동택배
' gsmnton     GSM NtoN
' daewoon 	대한글로벌
' direct 		배송지원
' korexg 		cj 대한통운국제특송
' cvsnet 		편의점택배

    select Case itenCode
        CASE "1" : TenDlvCode2MarketforDlvCode = "hanjin"     ''한진
        CASE "2" : TenDlvCode2MarketforDlvCode = "hyundai"     ''현대 -> 롯데택배로 변경됨 2017-03-13 김진영 수정
        CASE "3" : TenDlvCode2MarketforDlvCode = "korex"     ''대한통운 (CJ대한통운(CJGLS))
        CASE "4" : TenDlvCode2MarketforDlvCode = "korex"     ''CJ GLS (CJ대한통운(CJGLS))
        CASE "5" : TenDlvCode2MarketforDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2MarketforDlvCode = ""     ''삼성 HTH
        CASE "7" : TenDlvCode2MarketforDlvCode = "dongbu"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2MarketforDlvCode = "epost"     ''우체국택배
        CASE "9" : TenDlvCode2MarketforDlvCode = "kgbps"     ''KGB택배
        CASE "10" : TenDlvCode2MarketforDlvCode = ""     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2MarketforDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2MarketforDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2MarketforDlvCode = "yellow"     ''옐로우캡
        CASE "14" : TenDlvCode2MarketforDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2MarketforDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2MarketforDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2MarketforDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2MarketforDlvCode = "logen"     ''로젠택배
        CASE "19" : TenDlvCode2MarketforDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2MarketforDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2MarketforDlvCode = "kdexp"     ''경동택배
        CASE "22" : TenDlvCode2MarketforDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2MarketforDlvCode = ""     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2MarketforDlvCode = ""     ''사가와익스프레스
        CASE "25" : TenDlvCode2MarketforDlvCode = ""     ''하나로택배
        CASE "26" : TenDlvCode2MarketforDlvCode = "ilyang"     ''일양택배
        CASE "27" : TenDlvCode2MarketforDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2MarketforDlvCode = "dongbu"     ''동부익스프레스
        CASE "29" : TenDlvCode2MarketforDlvCode = ""     ''건영택배
        CASE "30" : TenDlvCode2MarketforDlvCode = ""     ''이노지스
        CASE "31" : TenDlvCode2MarketforDlvCode = "chunil"     ''천일택배
        CASE "33" : TenDlvCode2MarketforDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2MarketforDlvCode = "daesin"     ''대신화물택배
        CASE "35" : TenDlvCode2MarketforDlvCode = "cvsnet"     ''CVSnet택배  -
        CASE "37" : TenDlvCode2MarketforDlvCode = "hdexp"     ''합동택배  -
        CASE "38" : TenDlvCode2MarketforDlvCode = ""     ''GTX로지스
        CASE "39" : TenDlvCode2MarketforDlvCode = "dongbu"     ''KG로지스 - 동부익스프레스
        CASE "98" : TenDlvCode2MarketforDlvCode = "direct"     ''퀵서비스->직배송
        CASE "41" : TenDlvCode2MarketforDlvCode = "yellow"     ''드림택배(동부택배,옐로우캡)  ''2018/02/13
        CASE "42" : TenDlvCode2MarketforDlvCode = "cvsnet"     ''CU POST를 편의점택배로 해달라심..2019-03-08 김진영 수정
        CASE "99" : TenDlvCode2MarketforDlvCode = "ETC"     ''기타  0000033028
        CASE  Else
            TenDlvCode2MarketforDlvCode = "ETC"      ''기타발송
    end Select
end function

function TEN_TenDlvCode2CommonDlvCode(imallname,itenCode)
    if (LCASE(imallname)="lottecom") or (LCASE(imallname)="lottecomm") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2LotteDlvCode(itenCode)
    elseif (LCASE(imallname)="lotteimall") then
		If (Now() > #09/01/2015 00:00:00#) Then
			TEN_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallNewDlvCode(itenCode)
		Else
			TEN_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallDlvCode(itenCode)
		End If
    elseif (LCASE(imallname)="interpark") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2InterParkDlvCode(itenCode)
    elseif (LCASE(imallname)="cjmall") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2cjMallDlvCode(itenCode)
    elseif (LCASE(imallname)="gseshop") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2GSShopDlvCode(itenCode)
    elseif (LCASE(imallname)="homeplus") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2HomeplusDlvCode(itenCode)
    elseif (LCASE(imallname)="ezwel") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2EzwelDlvCode(itenCode)
    elseif (LCASE(imallname)="auction1010") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2AuctionDlvCode(itenCode)
    elseif (LCASE(imallname)="gmarket1010") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2GmarketDlvCode(itenCode)
    elseif (LCASE(imallname)="nvstorefarm") or (LCASE(imallname)="nvstoremoonbangu") or (LCASE(imallname)="Mylittlewhoopee") or (LCASE(imallname)="nvstoregift") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2NvstorefarmDlvCode(itenCode)
    elseif (LCASE(imallname)="11st1010") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode211stDlvCode(itenCode)
    elseif (LCASE(imallname)="ssg") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2SsgDlvCode(itenCode)
    elseif (LCASE(imallname)="halfclub") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2HalfClubDlvCode(itenCode)
    elseif (LCASE(imallname)="coupang") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2CoupangDlvCode(itenCode)
    elseif (LCASE(imallname)="hmall1010") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2HmallDlvCode(itenCode)
    elseif (LCASE(imallname)="wmp") OR (LCASE(imallname)="wmpfashion") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2WMPDlvCode(itenCode)
    elseif (LCASE(imallname)="lotteon") then
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2LotteonDlvCode(itenCode)
    elseif (LCASE(imallname)="gsisuper") or (LCASE(imallname)="lfmall") then
        '// gsisuper => 사방넷API 이용
        TEN_TenDlvCode2CommonDlvCode = TenDlvCode2SabangNetDlvCode(itenCode)
    end if
end function

function Fn_ActOutMall_CateSummary(iMallID)
    dim sqlStr
    sqlStr = "exec db_item.dbo.sp_Ten_OutMall_CateSummary '"&iMallID&"'"

    dbget.Execute sqlStr
end function

Function Fn_AcctFailTouch(iMallID, iitemid, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

	If (iMallID = "lotteCom") Then
		strSql = "Update R" & VBCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_lotte_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "lotteimall") Then
		strSql = "Update R"&VbCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_LTiMall_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "interpark") Then
		strSql = "Update R"&VbCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_item.dbo.tbl_interpark_reg_item R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
    ELSEIF (iMallID="gsshop") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_item.dbo.tbl_gsshop_regitem R"&VbCRLF
        strSql = strSql &" where itemid="&iitemid&VbCRLF
        dbget.Execute(strSql)
	ElseIf (iMallID = "cjmall") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_item.dbo.tbl_cjmall_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "ezwel") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_ezwel_regItem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "homeplus") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_homeplus_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "auction1010") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_auction_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "nvstorefarm") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_nvstorefarm_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "nvstorefarmclass") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_nvstorefarmclass_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "nvstoremoonbangu") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_nvstoremoonbangu_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "Mylittlewhoopee") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_Mylittlewhoopee_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "nvstoregift") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.dbo.tbl_nvstoregift_regitem as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "11stmy") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_my11st_regItem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "gmarket1010") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_gmarket_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "11st1010") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_11st_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "sabangnet") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_sabangnet_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "ssg") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_ssg_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	ElseIf (iMallID = "halfclub") Then
		strSql = ""
		strSql = strSql & "UPDATE R"&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_halfclub_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "coupang") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_coupang_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "hmall1010") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_hmall_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "WMP") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_wemake_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "wmpfashion") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = isnull(accFailCnt, 0) + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_wfwemake_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "lotteon") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_lotteon_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "lfmall") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_lfmall_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "shintvshopping") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_shintvshopping_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "wetoo1300k") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_wetoo1300k_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
  	ElseIf (iMallID = "skstoa") Then
		strSql = ""
		strSql = strSql & "UPDATE R "&VbCRLF
		strSql = strSql &" SET accFailCnt = accFailCnt + 1" & VBCRLF
		strSql = strSql &" ,lastErrStr = convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" FROM db_etcmall.[dbo].[tbl_skstoa_regitem] as R" & VBCRLF
		strSql = strSql &" WHERE itemid = "&iitemid & VBCRLF
		dbget.Execute(strSql)
	End If
End Function

Function Fn_AcctFailTouchOption(iMallID, iidx, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

    IF (iMallID="gsshop") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_etcmall.dbo.tbl_gsshopAddoption_regitem R"&VbCRLF
        strSql = strSql &" where midx="&iidx&VbCRLF
        dbget.Execute(strSql)
    ElseIF (iMallID="lotteimall") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_etcmall.dbo.tbl_ltimallAddOption_regItem R"&VbCRLF
        strSql = strSql &" where midx="&iidx&VbCRLF
        dbget.Execute(strSql)
    ElseIF (iMallID="lotteCom") THEN
        strSql = "Update R"&VbCRLF
        strSql = strSql &" SET accFailCnt=accFailCnt+1"&VbCRLF
        strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')"&VbCRLF
        strSql = strSql &" From db_etcmall.dbo.tbl_lotteAddOption_regItem R"&VbCRLF
        strSql = strSql &" where midx="&iidx&VbCRLF
        dbget.Execute(strSql)
	End If
End Function

Function Fn_AcctFailTouchEtcOption(iMallID, iitemid, iitemoption, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

	If (iMallID = "zilingo") Then
		strSql = "Update R" & VBCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_etcmall.dbo.tbl_zilingo_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		strSql = strSql &" and itemoption = '"&iitemoption&"' "
		dbget.Execute(strSql)
	End If
End Function

Function Fn_AcctFailTouchEtc(iMallID, iitemid, iLastErrStr)
	Dim strSql
	iLastErrStr = html2db(iLastErrStr)

	If (iMallID = "shopify") Then
		strSql = "Update R" & VBCRLF
		strSql = strSql &" SET accFailCnt=accFailCnt+1" & VBCRLF
		strSql = strSql &" ,lastErrStr=convert(varchar(100),'"&iLastErrStr&"')" & VBCRLF
		strSql = strSql &" From db_etcmall.dbo.tbl_shopify_regItem R" & VBCRLF
		strSql = strSql &" where itemid="&iitemid & VBCRLF
		dbget.Execute(strSql)
	End If
End Function

function Fn_AcctFailLog(iMallID,iitemid,ErrMsg,ErrCode)
    Dim sqlStr
    ''db_log.dbo.tbl_interparkEdit_log
    IF (iMallID="lotteCom") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(lotteGoodNo,lotteTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_lotte_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ELSEIF (iMallID="lotteimall") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.LTimallGoodno,R.LtiMallTmpGoodNo), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_ltimall_regItem R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr

    ELSEIF (iMallID="interpark") THEN
        sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
        sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
        sqlStr = sqlStr & " select R.itemid, isNULL(R.interparkprdno,''), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
        sqlStr = sqlStr & " ,'"&iMallID&"'" & VbCrlf
        sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
        sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
        sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
        sqlStr = sqlStr & " where R.itemid=" & iitemid & VbCrlf
        'rw sqlStr
        dbget.execute sqlStr
    ENd IF
end function

function Fn_AcctFailLogNone(iMallID,iitemid,ioutmallPrdno,ioutmallsellyn,ioutmallsellcash,ioutmallbuycash,ErrMsg,ErrCode)
    Dim sqlStr
    sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
    sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode, mallid)" & VbCrlf
    sqlStr = sqlStr & " values("&iitemid& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallPrdno&"'"& VbCrlf
    sqlStr = sqlStr & " ,"&ioutmallsellcash& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallbuycash&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&ioutmallsellyn&"'"& VbCrlf
    sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(ErrMsg)&"')"& VbCrlf
    sqlStr = sqlStr & " ,'"&ErrCode&"'"& VbCrlf
    sqlStr = sqlStr & " ,'"&iMallID&"')"& VbCrlf
    dbget.execute sqlStr
end function
%>
