<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/ezwel/ezwelItemcls.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
'// 2014-08-27, skyer9
''Server.ScriptTimeout = 60
'' response.write lotteAuthNo
'' response.end
Dim refer
refer = request.ServerVariables("HTTP_REFERER")

Dim sqlStr, buf
Dim i, j, k

'�ֹ��Ϸ� (1001) /����غ��� (1002) /����� (1003) /����Ϸ� (1004) /�ֹ���� (1005) /��ǰ��û (1007)
'��ǰ�Ϸ� (1008) /��ȯ��û (1011) /��ȯ�Ϸ� (1012) /��ǰ�� �ֹ���� (1009) /���� (1010)/ǰ����ҿ�û (1013)/ǰ����� (1014)

'// ============================================================================
'// [divcd]
'// ============================================================================
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�
'
'A000			�±�ȯ���
'A100			��ǰ���� �±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'A007			ī��,��ü,�޴�����ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)
'A012			�±�ȯ��ǰ(��ü���)

'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
'A112			��ǰ���� �±�ȯ��ǰ(��ü���)
'// ============================================================================

Dim mode
Dim sellsite
Dim reguserid
Dim AssignedRow
Dim ErrMsg

Dim resultCount

Dim divcd, yyyymmdd, idx, finUserid
Dim getDivCD, sDate, eDate

Dim postParam
Dim objXML, xmlDOM, strSql
Dim retCode, goodsCd, iMessage, oMsg, ocount, stdt, eddt
Dim parentNodes, parentSubNodes, Nodes, masterSubNodes
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
finUserid	= session("ssBctID")
If finUserid = "" Then
	finUserid = "system"
End If

If (mode = "getxsitecslist") Then
    If (sellsite="ezwel") Then
    	ErrMsg = ""
		getDivCD = Trim(application("xSiteGetEzwelCS_DIVCD"))
		If (getDivCD = "") Then
			getDivCD = "A008"
		ElseIf (getDivCD = "A004") Then
			getDivCD = "A008"
		Else
			getDivCD = "A004"
		End If

		postParam = "cspCd="&cspCd&"&crtCd="&crtCd
		If getDivCD = "A008" Then				'�ֹ����
			stdt = getLastCSInputDT("ordercancel")
		Else									'��ǰ
			stdt = getLastCSInputDT("return")
		End If
		eddt = Replace(Date,"-","") & Replace(FormatDateTime(Now,4),":","") & Right(Now,2)
		postParam = postParam & "&startDate="&stdt&"&endDate="&eddt

'		If Hour(Now()) < 6 then
'			postParam = postParam & "&startDate="&Replace(Date-1,"-","") & "000000"&"&endDate="&Replace(Date-1,"-","") & "235959"
'		Else
'			postParam = postParam & "&startDate="&Replace(Date,"-","") & "000000"&"&endDate="&Replace(Date,"-","") & Replace(FormatDateTime(Now,4),":","") & Right(Now,2)
'		End If

		If getDivCD = "A008" Then				'�ֹ����
			postParam = postParam & "&orderStatus=1005"
		Else									'��ǰ
			postParam = postParam & "&orderStatus=1007"
		End If

'		On Error Resume Next
		Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.open "POST", "http://api.ezwel.com/if/api/orderListAPI.ez", false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
			objXML.send(postParam)
			If objXML.Status = "200" Then
				Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
					If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
						'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
					End If
					retCode = xmlDOM.getElementsByTagName("resultCode").item(0).text
					If retCode = "200" Then		'����(200)
						Dim retVal, succCnt, failCnt
						Dim OutMallOrderSerial, OrgDetailKey, CSDetailKey, rcvrNm, rcvrTelNum, rcvrMobile, rcvrPost, rcvrAddr1, rcvrAddr2, orderDt, orderQty, sndNm, sndTelNum, sndMobile, orderReqContent
						succCnt = 0
						failCnt = 0
						Set parentNodes = xmlDOM.getElementsByTagName("arrOrderList")
							For each parentSubNodes in parentNodes
								OutMallOrderSerial = parentSubNodes.getElementsByTagName("orderNum").item(0).text				'�ֹ���ȣ
								CSDetailKey = parentSubNodes.getElementsByTagName("aspOrderNum").item(0).text					'(CS)�ֹ���ȣ
								sndNm = parentSubNodes.getElementsByTagName("sndNm").item(0).text								'�����ڸ�
								sndTelNum = parentSubNodes.getElementsByTagName("sndTelNum").item(0).text						'������ ��ȭ��ȣ
								sndMobile = parentSubNodes.getElementsByTagName("sndMobile").item(0).text						'������ �޴���
								rcvrNm =  parentSubNodes.getElementsByTagName("rcvrNm").item(0).text							'�����θ�
								rcvrTelNum = parentSubNodes.getElementsByTagName("rcvrTelNum").item(0).text						'������ ��ȭ��ȣ
								rcvrMobile = parentSubNodes.getElementsByTagName("rcvrMobile").item(0).text						'������ �޴���
								rcvrPost = parentSubNodes.getElementsByTagName("rcvrPost").item(0).text							'�����ȣ
								rcvrAddr1 = parentSubNodes.getElementsByTagName("rcvrAddr1").item(0).text						'�ּ�
								rcvrAddr2 = parentSubNodes.getElementsByTagName("rcvrAddr2").item(0).text						'���ּ�
								orderDt = LEFT(parentSubNodes.getElementsByTagName("orderDt").item(0).text, 8)					'�ֹ��� | ����Ͻú���(YYYYMMDDhh24miss)
								orderDt = LEFT(orderDt, 4) &"-"& MID(orderDt, 5,2) &"-"& right(orderDt,2)
'								rw "�������� : " & parentSubNodes.getElementsByTagName("dlvrHopeDt").item(0).text				'�������� | ����Ͻú���(YYYYMMDDhh24miss)
								orderReqContent = parentSubNodes.getElementsByTagName("orderReqContent").item(0).text			'��ۿ�û����
'								rw "###############################################################"
								strSql = " select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = 'ezwel' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
								rsget.CursorLocation = adUseClient
								rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
								If (Not rsget.Eof) then
									retVal = false
								Else
									retVal = true
								End if
								rsget.Close

								Set Nodes = parentSubNodes.getElementsByTagName("arrOrderGoods")
									For each masterSubNodes in Nodes
										OrgDetailKey = masterSubNodes.getElementsByTagName("orderGoodsNum")(0).Text				'�ֹ����� / ��ٱ��� ����
'										rw "����� ID : " & masterSubNodes.getElementsByTagName("cspDlvrId")(0).Text			'����� ID
'										rw "CP��ü �ڵ� : " & masterSubNodes.getElementsByTagName("cspCd")(0).Text				'CP��ü �ڵ�
'										rw "��ǰ�ڵ� : " & masterSubNodes.getElementsByTagName("goodsCd")(0).Text				'��ǰ�ڵ�
'										rw "��ü��ǰ�ڵ� : " & masterSubNodes.getElementsByTagName("cspGoodsCd")(0).Text		'��ü��ǰ�ڵ�
'										rw "��ǰ�� : " & masterSubNodes.getElementsByTagName("goodsNm")(0).Text					'��ǰ��
										orderQty = masterSubNodes.getElementsByTagName("orderQty")(0).Text						'�ֹ�����
'										rw "��ǰ�ɼ� : " & masterSubNodes.getElementsByTagName("optionContent")(0).Text			'��ǰ�ɼ�(������ ^) Ex) ���� �� �뷮 ����:500GB ����^'�߰� ������ǰ:���þ���^
'										rw "���԰� : " & masterSubNodes.getElementsByTagName("buyPrice")(0).Text				'���ּ� | ���԰�
'										rw "�ǸŰ� : " & masterSubNodes.getElementsByTagName("salePrice")(0).Text				'�ǸŰ� | �ɼǰ�������
'										rw "���ξ� : " & masterSubNodes.getElementsByTagName("dccpnPrice")(0).Text				'���ξ� | ������������εȱݾ�
'										rw "�ù� ����� ��ȣ : " & masterSubNodes.getElementsByTagName("dlvrNo")(0).Text		'�ù� ����� ��ȣ
'										rw "��۾�ü �ڵ� : " & masterSubNodes.getElementsByTagName("dlvrCd")(0).Text			'��۾�ü �ڵ� | ����÷��
'										rw "����Ͻ� : " & masterSubNodes.getElementsByTagName("dlvrDt")(0).Text				'����Ͻ� | ����Ͻú���(YYYYMMDDhh24miss)
'										rw "�ɼǰ��� : " & masterSubNodes.getElementsByTagName("optionAddPrice")(0).Text		'�ɼǰ���
'										rw "�ֹ����� : " & masterSubNodes.getElementsByTagName("orderStatus")(0).Text			'�ֹ�����
'										rw "��ۺ� ������� �ڵ� : " & masterSubNodes.getElementsByTagName("dlvrPayCd")(0).Text	'��ۺ� ������� �ڵ� | ������(1001), ����(1002)
'										rw "��ۺ� : " & masterSubNodes.getElementsByTagName("dlvrPrice")(0).Text				'��ۺ�
'										rw "��ۿϷ��� : " & masterSubNodes.getElementsByTagName("dlvrFinishDt")(0).Text		'��ۿϷ��� | �����(YYYYMMDD)
'										cancelDt = masterSubNodes.getElementsByTagName("cancelDt")(0).Text						'�ֹ������ | ����Ͻú���(YYYYMMDDhh24miss)
'										rw "======================================================"
										strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = 'ezwel' and OutMallOrderSerial = '" & CStr(OutMallOrderSerial) & "' and OrgDetailKey = '" & CStr(OrgDetailKey) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "') "
										strSql = strSql & " BEGIN "
										strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
										strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
										strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
										strSql = strSql & " 	('" & CStr(getDivCD) & "', '�ܼ�����', 'ezwel', '" & html2db(CStr(OutMallOrderSerial)) & "', '"&CStr(sndNm)&"', '', '"&html2db(CStr(sndTelNum))&"', '"&html2db(CStr(sndMobile))&"', '" & html2db(CStr(rcvrNm)) & "', "
										strSql = strSql & "		'" & html2db(CStr(rcvrTelNum)) & "', '" & html2db(CStr(rcvrMobile)) & "', '" & html2db(CStr(rcvrPost)) & "', '" & html2db(CStr(rcvrAddr1)) & "', '" & html2db(CStr(rcvrAddr2)) & "', '"&html2db(CStr(orderReqContent))&"' "
										strSql = strSql & "		, '" & html2db(CStr(orderDt)) & "', '" & html2db(CStr(OrgDetailKey)) & "', '" & html2db(CStr(CSDetailKey)) & "', " & CStr(orderQty) & ") "
										strSql = strSql & " END "
										''rw strSql
										dbget.Execute(strSql)
				                        If (retVal) Then
				                            succCnt = succCnt + 1
				                        Else
											failCnt = failCnt + 1
				                        End If
									Next
								Set Nodes = nothing
							Next
						Set parentNodes = nothing
					    rw succCnt & "�� �Է�"
					    rw failCnt & "�� ����"

						If (succCnt > 0) then
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " , c.OrderName = o.OrderName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.sellsite = 'ezwel' "
							''rw strSql
							dbget.Execute(strSql)

							If getDivCD = "A008" Then
								strSql = " update c "
								strSql = strSql + " set c.currstate = 'B007' "
								strSql = strSql + " from "
								strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
								strSql = strSql + " left join db_temp.dbo.tbl_xSite_TMPOrder o "
								strSql = strSql + " on "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.SellSite = o.SellSite "
								strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
								strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
								strSql = strSql + " where "
								strSql = strSql + " 	1 = 1 "
								strSql = strSql + " 	and c.orderserial is NULL "
								strSql = strSql + " 	and o.SellSite is NULL "
								strSql = strSql + " 	and c.sellsite = 'ezwel' "
								strSql = strSql + " 	and c.currstate = 'B001' "
								strSql = strSql + " 	and c.divcd = 'A008' "
								''rw strSql
								dbget.Execute(strSql)
							end if
						End If

						If getDivCD = "A008" Then				'�ֹ����
							Call UpdateLastCSInputDT("ordercancel", date())
						Else									'��ǰ
							Call UpdateLastCSInputDT("return", date())
						End If

						If (getDivCD <> Trim(application("xSiteGetEzwelCS_DIVCD"))) then
							application("xSiteGetEzwelCS_DIVCD") = getDivCD
						End If
					End If
				Set objXML = Nothing
			End If
		Set xmlDOM = Nothing
'		On Error Goto 0
    End If
End If

Function getLastCSInputDT(mode)
	Dim sqlStr
	sqlStr = "select top 1 convert(varchar(10),LastCheckDate,21) as lastCSInputDt"
	sqlStr = sqlStr&" from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr&" where sellsite = 'ezwel' and csGubun = '" & CStr(mode) & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If (Not rsget.Eof) Then
		getLastCSInputDT = replace(rsget("lastCSInputDt"), "-", "") & "000000"
	Else
		getLastCSInputDT = "20161020000000"
	End If
	rsget.Close
End Function

Function UpdateLastCSInputDT(mode, dt)
	Dim sqlStr
	sqlStr = " UPDATE db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	sqlStr = sqlStr & " SET LastCheckDate = '" & CStr(dt) & "' "
	sqlStr = sqlStr & " WHERE sellsite = 'ezwel' and csGubun = '" & CStr(mode) & "' "
	dbget.Execute sqlStr
End Function
%>
<% rw "OK" %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
