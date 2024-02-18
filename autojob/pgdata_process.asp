<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<%
'###########################################################
' Description : PG����γ���
' Hieditor : 2011.04.22 �̻� ����
'			 2023.03.28 �ѿ�� ����(Apple Pay�߰�)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
'' /wAPI/autojob/pgdata_process.asp  '' �ǵ��� �������� �̰���
' �ݵ�� 3������ �ҽ��� �����ؾ� �մϴ�. �Ѱ��� ��ĥ��� ������ �ΰ��� ������ �ּ���.
' scm\admin\maechul\pgdata\pgdata_process.asp
' webadmin\admin\maechul\pgdata\pgdata_process.asp
' wapi\autojob\pgdata_process.asp

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.83","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write "inValid "
    dbget.Close()
    response.end
end if



Dim StopWatch(19)

sub StartTimer(x)
	StopWatch(x) = Timer
end Sub

function StopTimer(x)
	Dim EndTime

	EndTime = Timer

	if EndTime < StopWatch(x) Then
		EndTime = EndTime + (86400)
	end if

	StopTimer = EndTime - StopWatch(x)
end function

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, reguserid
dim logidx, orderno, matchidx
dim IsMatched

dim objData, objXML, xmlURL, objLine, xmlURLArr
dim PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate
dim appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid
dim lastipkumdate, searchipkumdate
dim prevPGkey, prevPrevPGkey, prevAppDivCode, prevPrevAppDivCode, IsDuplicate
dim yyyymmdd
dim subPgKey
dim tmpStr, arrOrderSerial, orderserial
dim reasonGubun
dim searchipkumdateMAX, force

dim objFSO, objOpenedFile
Dim targetFileName
Const ForReading = 1

dim yyyymm

mode = requestCheckvar(request("mode"),64)
logidx = requestCheckvar(request("logidx"),32)
orderno = requestCheckvar(request("orderno"),32)
yyyymmdd = requestCheckvar(request("yyyymmdd"),32)
reasonGubun = requestCheckvar(request("reasonGubun"),32)

yyyymm = requestCheckvar(request("yyyymm"),7)

reguserid = session("ssBctId")

dim sqlStr
dim i, j, k



if (mode="matchoneorder") then

    sqlStr = " select isNULL(orderserial,'') as orderserial " & VbCRLF
    sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_cardApp_log " & VbCRLF
    sqlStr = sqlStr & " where idx="&logidx&VbCRLF

	IsMatched = True

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    IsMatched = Not (rsget("orderserial") = "")
	end if
	rsget.Close

	if IsMatched then
		response.write "<script>alert('�̹� ��Ī�� �����Դϴ�.');</script>"
		response.write "�̹� ��Ī�� �����Դϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = m.idx, l.orderserial = m.orderno, l.shopid = m.shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master m "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and m.orderno = '" + CStr(orderno) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="matchsumipkum") then

	arrOrderSerial = Split(requestCheckvar(request("arrOrderSerial"),512), vbCrLf)
	tmpStr = ""

	for each orderserial in arrOrderSerial
		if (Len(orderserial) > 0) then
			if (Len(orderserial) <> 11) then
				response.write "<script>alert('�߸��� �ֹ���ȣ�Դϴ�.');</script>"
				response.write "�߸��� �ֹ���ȣ�Դϴ�." & orderserial
				dbget.close()
				response.end
			end if

			if (tmpStr = "") then
				tmpStr = " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			else
				tmpStr = tmpStr + " union all " & vbCrLf & " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			end if
		end if
	next

	if (tmpStr = "") then
		response.write "<script>alert('�Էµ� �ֹ���ȣ�� �����ϴ�.');</script>"
		response.write "�Էµ� �ֹ���ȣ�� �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, PGuserid, orgPayDate, PGmeachulDate) " & vbCrLf
	sqlStr = sqlStr + " select l.PGgubun, l.PGkey, T.orderserial, l.sitename, l.appDivCode, l.appMethod, l.appDate, l.cancelDate, 0, 0, 0, 0, l.ipkumdate, T.orderserial, l.PGuserid, l.orgPayDate, l.PGmeachulDate " & vbCrLf
	sqlStr = sqlStr + " from " & vbCrLf
	sqlStr = sqlStr + "	db_order.dbo.tbl_onlineApp_log l " & vbCrLf
	sqlStr = sqlStr + "	join ( " & vbCrLf

	sqlStr = sqlStr + tmpStr

	sqlStr = sqlStr + "	) T " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	left join db_order.dbo.tbl_onlineApp_log l2 " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "		and l.pggubun = l2.pggubun " & vbCrLf
	sqlStr = sqlStr + "		and l.pgkey = l2.pgkey " & vbCrLf
	sqlStr = sqlStr + "		and T.orderserial = l2.pgcskey " & vbCrLf
	sqlStr = sqlStr + "where " & vbCrLf
	sqlStr = sqlStr + "	1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	and l.pggubun = 'bankipkum' " & vbCrLf
	sqlStr = sqlStr + "	and l.appDivCode = 'A' " & vbCrLf
	sqlStr = sqlStr + "	and l.idx = " + CStr(logidx) + " " & vbCrLf
	''sqlStr = sqlStr + "	and l.PGCSkey = '' " & vbCrLf
	sqlStr = sqlStr + "	and l2.idx is NULL " & vbCrLf
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="regReasonGubun") then

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = '" + CStr(reasonGubun) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and IsNull(reasonGubun, '') not in ('030', '950') "
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="delmatchone") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = NULL, l.orderserial = NULL "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	''sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is not NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('�����Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancel") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopjumun_cardApp_log c "
	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopjumun_cardApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.cardAppNo = a.cardAppNo "
	''sqlStr = sqlStr + " 		and convert(VARCHAR(10), c.appDate, 127) = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " 		and DateDiff(d,a.appDate,c.appDate) < 1 "
	''sqlStr = sqlStr + " 		and c.shopid = a.shopid "
	sqlStr = sqlStr + " 		and ((c.shopid = a.shopid) or (a.shopid is NULL and c.cardReaderID = a.cardReaderID)) "
	sqlStr = sqlStr + " 		and c.cardPrice*-1 = a.cardPrice "
	sqlStr = sqlStr + " 		and c.appDivCode in ('C','P') "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and a.orderserial is NULL "
	sqlStr = sqlStr + " 	and c.orderserial is NULL "

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[0].');</script>"
		response.write "��Ī������ �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_shop.dbo.tbl_shopjumun_cardApp_log "
	sqlStr = sqlStr + " set shopJumunMasterIdx = -1, orderserial = '��Ҹ�Ī' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="addActLog") then

	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.pggubun = a.pggubun "
	sqlStr = sqlStr + " 	and o.pgkey = a.pgkey "
	sqlStr = sqlStr + " 	and o.pgcskey = Left(a.pgcskey, len(o.pgcskey)) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and o.idx = " + CStr(logidx) + " "

	PGCSkey = ""

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    PGCSkey = "-" + Format00(3, rsget("cnt"))
	end if
	rsget.Close

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select top 1 t.PGgubun, t.PGkey, t.PGCSkey + '" + CStr(PGCSkey) + "', t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, 0, 0, 0, 0, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and t.appPrice <> 0 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	response.write "<script>alert('�߰��Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnline") then

	PGkey = requestCheckvar(request("PGkey"),64)
	force = requestCheckvar(request("force"),1)

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	if (force = "Y") then
		sqlStr = sqlStr + " 	and c.pgkey = '" & PGkey & "' "
	else
		sqlStr = sqlStr + " 	and (convert(VARCHAR(10), IsNull(c.appDate,c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) or a.pggubun = 'bankipkum') "		'// �������ڿ� ������ڰ� �ٸ� ���, �ּ�ó�� �� ��Ī�Ѵ�.
	end if
	sqlStr = sqlStr + " 	and IsNull(c.sitename, '') = IsNull(a.sitename, '') "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " and a.orderserial is NULL "
	sqlStr = sqlStr + " and c.orderserial is NULL "
	''rw sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[1]. �������ڿ� ������ڰ� �ٸ� ��� �����ּ���.');</script>"
		response.write "��Ī������ �����ϴ�. �������ڿ� ������ڰ� �ٸ� ��� �����ּ���."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orderserial = '��Ҹ�Ī' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1 "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnlineDup") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	''sqlStr = sqlStr + " 	and convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 1 "
	sqlStr = sqlStr + " 	and c.sitename = a.sitename "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " and c.csasid is NULL "
	sqlStr = sqlStr + " and c.orderserial is NULL "		'// �ֹ���ȣ ���� ���
	sqlStr = sqlStr + " and c.idx > a.idx "
	''response.write sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		'// �ֹ���ȣ �ִ� ���
		sqlStr = " select top 1 a.idx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
		sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 15 "
		sqlStr = sqlStr + " 	and c.sitename = a.sitename "
		sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
		sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 1 = 1 "
		sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
		sqlStr = sqlStr + " and c.csasid is NULL "
		sqlStr = sqlStr + " and c.orderserial is not NULL "
		sqlStr = sqlStr + " and c.idx > a.idx "
		sqlStr = sqlStr + " and a.orderserial = c.orderserial "

		''response.write sqlStr
		matchidx = -1

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			matchidx = rsget("idx")
		end if
		rsget.Close
	end if

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[2].');</script>"
		response.write "��Ī������ �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1, reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(matchidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="getonpgdata") then

	'// ========================================================================
	'// INICIS
	if (yyyymmdd = "") then
		searchipkumdateMAX = ""
		sqlStr = " exec [db_cs].[dbo].[usp_getDayPlusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly
		if Not rsget.Eof then
			'// �ٹ��ϼ� ���� D+4 ��
			searchipkumdateMAX = rsget("plusworkday")
		end if
		rsget.close

		''rw searchipkumdateMAX
		''response.end

		lastipkumdate = searchipkumdateMAX

		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)
		ipkumdate = Replace(searchipkumdate, "-", "")

		'// ========================================================================
		'// �¶��� ���� ���곻��
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=teenxteen4&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=32"
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt04GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt14GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		''response.write xmlURL
		''response.end

		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 1000, 5 * 1000, 15 * 1000, 45 * 1000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			''response.write objXML.ResponseBody
			''response.end
			if (Trim(objXML.ResponseBody)<>"") then
				objData = BinaryToText(objXML.ResponseBody, "euc-kr")
			else
				if  (Not IsAutoScript) then
					response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[b]');</script>"
				end if
				response.write "������ ����Ÿ�� �����ϴ�[b]"
				dbget.close()
				response.end
			end if
		end if

		''response.write objXML.Status

		Set objXML  = Nothing

		if (InStr(objData, "NO DATA") > 0) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[1]"
			dbget.close()
			response.end
		end if

		lastipkumdate = searchipkumdate
	else
		lastipkumdate = yyyymmdd

		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)

		ipkumdate = Replace(searchipkumdate, "-", "")

		'// ========================================================================
		'// �¶��� ���� ���곻��
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=teenxteen4&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=32"
		''xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt04GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
		xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt14GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"

		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 1000, 5 * 1000, 15 * 1000, 45 * 1000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		Set objXML  = Nothing

		if (InStr(objData, "NO DATA") > 0) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[1]"
			response.write objData
			response.end
		end if
	end if
	''response.write objData
	''response.end

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				''sitename = "fingers"
                sitename = "wholesale"					'// 2022-04-21
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// �κ����
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

			sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
			sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

		end if
	next

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�Ա����� : " + CStr(searchipkumdate) + "');</script>"
	end if

elseif (mode="getpaycoT") Then

	'// ========================================================================
	'// ������ ���γ���
	'// ========================================================================

	''yyyymmdd = "2017-06-11"

	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateAdd("d", -1, Now()),10)
	End If

	'// ���� : https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// �׼� : https://dev-apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// CSV ������ Response ����
	'// ?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101
	'// ?serviceCode=ST_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101

	ReDim xmlURLArr(2)
	xmlURLArr(0) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(1) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=RR0VR3&token=RR0VR3-8EA5C0D-768CA-5F33225&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(2) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=8973MQ&token=8973MQ-5CBF5E4-7B1A9-D8FD548&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")

	objData = ""

	For Each xmlURL In xmlURLArr
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 1000, 5 * 1000, 15 * 1000, 45 * 1000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" And Len(objXML.ResponseText) > 0 Then
			objData = objData & vbLf & BinaryToText(objXML.ResponseBody, "UTF-8")
		else
		    response.write "NODATA:"&xmlURL
		end if

		Set objXML  = Nothing
	Next

	''response.write objData
	''response.end

	if (objData = "") then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		dbget.close()
		response.end
	end If

	''response.Write objData

	objData = Split(objData, vbLf)

	''response.Write UBound(objData)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, Chr(9))		'// �ǹ���

		If (UBound(objLine) > 0) Then
			If (IsNumeric(objLine(0))) Then
				''response.Write objData(i) & "<br />"


				PGgubun			= "payco"
				PGuserid 		= "payco"			'// PGuserid, sitename �� tbl_order_PaymentEtc ���� �����;� ��
				sitename 		= "10x10"

				'// ���� : �ſ�ī��/������ ����/PAYCO ����Ʈ �� �ɰ����� ���´�. ���������� ���ľ� �Ѵ�.
				if (objLine(7) = "����") then
					'// ==============================
					PGkey		= objLine(1)
					appDivCode	= "A"
					PGCSkey		= ""

					appDate		= objLine(0)
					cancelDate	= "NULL"
				else
					'// ==============================
					'// �κ����(���/�κ���Ҵ� ���γ������� �ݾ׺񱳷� ã�ƾ� �Ѵ�.)
					PGkey		= objLine(1)
					appDivCode	= "R"
					PGCSkey		= objLine(3)

					appDate		= "NULL"
					cancelDate	= objLine(0)
				end If

				appMethod = "100"			'// �ſ�ī�常 �ִ�.

				appPrice		= objLine(5)
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0

				ipkumdate		= ""

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end If

				sqlStr = " if exists( "
				sqlStr = sqlStr + " 	select 1 "
				sqlStr = sqlStr + " 	from db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " ) "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	update db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	set appPrice = appPrice + '" & appPrice & "' "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " End "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr

			End If
		End If
	Next

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'payco' "
	sqlStr = sqlStr + " 	and t.appDivCode = 'A' "				'// ���γ�����
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.appDivCode = 'C', r.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.PGgubun = r.PGgubun "
	sqlStr = sqlStr + " 		and a.PGkey = r.PGkey "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " 		and r.appDivCode <> 'A' "
	sqlStr = sqlStr + " 		and a.appPrice = r.appPrice*-1 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	a.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and ((l.PGCSkey = t.PGCSkey) or (l.PGCSkey = 'CANCELALL')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'payco' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + " [9]');</script>"
	end If


elseif (mode="getpaycoS") Then

	'// ========================================================================
	'// ������ ���곻��
	'// ========================================================================

	''yyyymmdd = "2017-06-13" ''�ּ�ó��..;;

	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateAdd("d", -2, Now()),10)   ''2016/12/23 d-2�� ���� ���� 4�ÿ� ������ ���µ���.
	End If

	'// ���� : https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// �׼� : https://dev-apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// CSV ������ Response ����
	'// ?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101
	'// ?serviceCode=ST_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101

	ReDim xmlURLArr(2)
	xmlURLArr(0) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(1) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=RR0VR3&token=RR0VR3-8EA5C0D-768CA-5F33225&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(2) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=8973MQ&token=8973MQ-5CBF5E4-7B1A9-D8FD548&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")


	For Each xmlURL In xmlURLArr
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 1000, 5 * 1000, 15 * 1000, 45 * 1000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" And Len(objXML.ResponseText) > 0 Then
			objData = objData & vbLf & BinaryToText(objXML.ResponseBody, "UTF-8")
		else
		    response.write "NODATA:"&xmlURL
		end if

		Set objXML  = Nothing
	Next

	if (objData = "") then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		dbget.close()
		response.end
	end If

	''response.Write objData
	''response.End

	objData = Split(objData, vbLf)


	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, Chr(9))		'// �ǹ���

		If (UBound(objLine) > 0) Then
			If (IsNumeric(objLine(0))) Then
				''response.Write objData(i) & "<br />"

				PGgubun			= "payco"
				PGuserid 		= "payco"			'// PGuserid, sitename �� tbl_order_PaymentEtc ���� �����;� ��
				sitename 		= "10x10"

				'// ���� : �ſ�ī��/������ ����/PAYCO ����Ʈ �� �ɰ����� ���´�. ���������� ���ľ� �Ѵ�.
				if (objLine(14) = "����") then
					'// ==============================
					PGkey		= objLine(10)
					appDivCode	= "A"
					PGCSkey		= ""

					appDate		= objLine(1)
					cancelDate	= "NULL"
				else
					'// ==============================
					'// �κ����(���/�κ���Ҵ� ���γ������� �ݾ׺񱳷� ã�ƾ� �Ѵ�.)
					PGkey		= objLine(10)
					appDivCode	= "R"
					PGCSkey		= objLine(12)

					appDate		= "NULL"
					cancelDate	= objLine(1)
				end If

				appMethod = "100"			'// �ſ�ī�常 �ִ�.

				appPrice		= objLine(16)
				commPrice		= objLine(17)
				commVatPrice	= objLine(20)
				jungsanPrice	= objLine(21)

				ipkumdate		= objLine(0)

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end If

				sqlStr = " if exists( "
				sqlStr = sqlStr + " 	select 1 "
				sqlStr = sqlStr + " 	from db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " ) "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	update db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	set appPrice = appPrice + '" & appPrice & "', commPrice = commPrice + '" & commPrice & "', commVatPrice = commVatPrice + '" & commVatPrice & "', jungsanPrice = jungsanPrice + '" & jungsanPrice & "' "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " End "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr

			End If
		End If
	Next

	'// ������ ���곻�� ��, �κ���� ������ �Ǿ� ��ü��ҵǸ� ��ü��� ���곻�� �ѰǸ� �´�.
	sqlStr = " update r "
	sqlStr = sqlStr + " set r.appDivCode = 'C', r.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.PGgubun = r.PGgubun "
	sqlStr = sqlStr + " 		and a.PGkey = r.PGkey "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " 		and r.appDivCode <> 'A' "
	sqlStr = sqlStr + " 		and a.appPrice = r.appPrice*-1 "
	''sqlStr = sqlStr + " 		and IsNull(a.cancelDate,a.appDate) = IsNull(r.cancelDate,r.appDate) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	a.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.commPrice = t.commPrice*-1, l.commVatPrice = t.commVatPrice*-1, l.jungsanPrice = t.jungsanPrice, l.ipkumdate = t.ipkumdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 	and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 	and ((l.appDivCode = t.appDivCode) or (l.appDivCode = 'R' and t.appDivCode = 'C')) " '// ���� : 17020632889
	sqlStr = sqlStr + " 	and IsNull(l.cancelDate,l.appDate) = IsNull(t.cancelDate,t.appDate) "
	''sqlStr = sqlStr + " 	and l.appPrice = t.appPrice "			'// �ݾ��� �޶� �Է��Ѵ�.
	sqlStr = sqlStr + " where t.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('�������� : " + CStr(yyyymmdd) + " [9]');</script>"
		dbget.close()
		response.end
	end If

	''response.Write "aaa"
	''response.end

elseif (mode="getonpgdatahppre") then

	'// ========================================================================
	'// INICIS �ڵ���(�����۾�)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('OK');</script>"
	end If

elseif (mode="getonpgdatahp") Then

	Call StartTimer(0)

	'// ========================================================================
	'// INICIS �ڵ���
	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateSerial(Year(Now()), Month(Now())+2, Day(Now()) - 2), 10)
	end If

	ipkumdate = Replace(yyyymmdd, "-", "")

	xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc/UrlSendExtraDc.jsp?urlid=teenteen10&passwd=cube1010??&date=" & ipkumdate & "&flgdate=P"

	objData = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

	objXML.setTimeouts 5 * 1000, 5 * 1000, 90 * 1000, 90 * 1000
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
	end if

	Set objXML  = Nothing

	if (InStr(objData, "NO DATA") > 0) then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		response.write objData
		response.end
	end if

	''Response.Write "Elapsed time was: " & StopTimer(0)
	''dbget.Close()
	''Response.End

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	sqlStr = ""
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				''sitename = "fingers"
                sitename = "wholesale"					'// 2022-04-21
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
			elseif (objLine(4) = "teenteen10") then
				if (Left(objLine(8),6) = "INIMX_") Then
					sitename = "10x10mobile"
				Else
					sitename = "10x10"
				End If
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// �κ����
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			elseif (objLine(3) = "MO") then
				appMethod = "400"
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)


			If (sqlStr = "") Then
				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			Else
				sqlStr = sqlStr + ", ('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			End If

			If (i <> 0) And ((i mod 500) = 0) Then
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr

				sqlStr = ""
			End If
		end if
	Next

	If (sqlStr <> "") Then
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		sqlStr = ""
	End If

	''rw "aaa" & Now()
	''dbget.close()
	''response.end

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�Ա����� : " + CStr(yyyymmdd) + " [" & StopTimer(0) & " sec]');</script>"
	end If

elseif (mode="getonpgdatakakaopayT") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	''yyyymmdd = "20170309"

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	''yyyymmdd = "20150819"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''response.write targetFileName
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A :����, C : ���(��ü��� or �κ����)
				Select Case objLine(3)
					Case "A"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "A"
						PGCSkey		= ""

						appDate		= objLine(2)
						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "C"
						PGCSkey		= "UNKNOWN"

						appDate		= "NULL"
						cancelDate		= objLine(2)
					Case Else
						'// ==============================
						PGkey		= objLine(5)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(8)
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0

				If appDivCode <> "A" Then
					appPrice = appPrice * -1
				End If

				ipkumdate		= ""

				'// 20130503
				'// (2013-05-03)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				''ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		if application("Svr_Info") <> "Dev" Then
			'// �׼��� ����Ÿ �����Ƿ� ���۵�

			'// ��ü��� or �κ����
			sqlStr = " update T "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " T.PGCSkey = (case when l.clogIdx is NULL then 'CANCELALL' else T.pgkey end) "
			sqlStr = sqlStr + " , T.appDivCode = (case when l.clogIdx is NULL then 'C' else 'R' End) "
			sqlStr = sqlStr + " , T.orderserial = (case when l.clogIdx is NULL then NULL else l.orderserial End) "
			sqlStr = sqlStr + " , T.cancelDate = (case when l.clogIdx is NULL then T.cancelDate else l.regdate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_card_cancel_log] l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.pgkey = l.newtid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.pggubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 	and T.PGCSkey = 'UNKNOWN' "
			dbget.execute sqlStr

			'// �ֹ���ȣ, ��������
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and o.ipkumdiv > 3 "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// ���ų���
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// ��ü�������
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.cancelDate = a.finishdate "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and T.orderserial = a.orderserial "
			sqlStr = sqlStr + " 		and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 		and a.divcd = 'A007' "
			sqlStr = sqlStr + " 		and a.currstate = 'B007' "
			sqlStr = sqlStr + " 		and a.deleteyn <> 'Y' "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

		End If

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		' dbget.Close()
		' Response.end

		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ���� ������ �����ϴ�.[0]');</script>"
		end if
		response.write "�ŷ���� ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatakakaopayS") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mS20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gS20150818.csv

	''yyyymmdd = "20170309"

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	''yyyymmdd = "20150827"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gS" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			''rw objLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : ����, C : ���, P: �κ����
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// �κ����
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				If (appDivCode <> "A") Then
					appPrice = appPrice * -1
				End If

				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))

				commPrice = commPrice - commVatPrice

				If (appDivCode = "A") Then
					commPrice = commPrice * -1
					commVatPrice = commVatPrice * -1
				End If

				jungsanPrice	= appPrice + (commPrice + commVatPrice)

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") Then
					If (appDate = "") Then
						appDate = "NULL"
					Else
						appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
					End If
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " update l "
		sqlStr = sqlStr + " set l.commPrice = T.commPrice, l.commVatPrice = T.commVatPrice, l.jungsanPrice = T.jungsanPrice, l.ipkumdate = T.ipkumdate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and T.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 		and T.PGgubun = l.PGgubun "
		sqlStr = sqlStr + " 		and T.PGkey = l.PGkey "
		sqlStr = sqlStr + " 		and T.appDivCode = l.appDivCode "
		sqlStr = sqlStr + " 		and T.PGCSkey = l.PGCSkey "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		'// ���� ��ü��Ҵ� ������ �ȿ´�.
		sqlStr = " update db_order.dbo.tbl_onlineApp_log "
		sqlStr = sqlStr + " set jungsanPrice = appPrice, ipkumdate = convert(varchar(10), IsNull(cancelDate,appDate), 127) "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 	and PGkey in ( "
		sqlStr = sqlStr + " 		select a.PGkey "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " 			join db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " 			on "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and a.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 				and a.PGgubun = c.PGgubun "
		sqlStr = sqlStr + " 				and a.PGkey = c.PGkey "
		sqlStr = sqlStr + " 				and a.appDivCode = 'A' "
		sqlStr = sqlStr + " 				and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 				and a.PGCSkey = '' "
		sqlStr = sqlStr + " 				and c.PGCSkey = 'CANCELALL' "
		sqlStr = sqlStr + " 				and convert(varchar(10), a.appDate, 127) = convert(varchar(10), c.cancelDate, 127) "
		sqlStr = sqlStr + " 				and a.ipkumdate = '' "
		sqlStr = sqlStr + " 				and a.ipkumdate = c.ipkumdate "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and ipkumdate = '' "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('�������� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ������ �����ϴ�.[0]');</script>"
		end if
		response.write "������ ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

	''dbget.Close
	''response.end

elseif (mode="getonpgdatakakaopay") then
	'// ========================================================================
	'// īī��PAY

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : ����, C : ���, P: �κ����
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// �κ����
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))
				jungsanPrice	= appPrice - commPrice

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ������ �����ϴ�.[0]');</script>"
		end if
		response.write "������ ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatauplus") then

	'// ========================================================================
	'// UPLUS

	'// ����/�������
	 ''yyyymmdd = "2017-10-30"

	if (yyyymmdd = "") then
		lastipkumdate = "2012-12-31"

		'// ��������
		sqlStr = " select max(PGmeachulDate) as lastipkumdate " & VbCRLF
		sqlStr = sqlStr & " from db_order.dbo.tbl_onlineApp_log " & VbCRLF
		sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
		''response.write sqlStr

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			if Not IsNull(rsget("lastipkumdate")) then
				lastipkumdate = rsget("lastipkumdate")
			end if
		end if
		rsget.Close

		''lastipkumdate = "2017-10-01"

		for i = 0 to 20
			'// TODO : 20�� �̻� �Աݾ��� ������ ����
			searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)) + 1)), 10)

			if False and (searchipkumdate >= Left(now, 10)) then
				if  (Not IsAutoScript) then
					response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[" & i & "]');</script>"
				end if
				response.write "������ ����Ÿ�� �����ϴ�[00]" & searchipkumdate
				response.end
			end if

			ipkumdate = Replace(searchipkumdate, "-", "")

			'// ========================================================================
			'// �¶��� ���� ���곻�� 99�� ��������.
			response.write "������"&CStr(ipkumdate) & "<br />"
			xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(ipkumdate) + "&key=2beec91670e1f2840a7ac80adde00e49"

			objData = ""

			Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

			objXML.Open "GET", xmlURL, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			if (request.serverVariables("REMOTE_ADDR")="110.93.128.113") then  ''�߰� ��ġ�۾��� timeout �ø�..
			    objXML.setTimeouts 30000,60000,60000,60000 ''2016/08/21 �߰�
		    end if
			objXML.Send()

			if objXML.Status = "200" then
			    if (Trim(objXML.ResponseBody)<>"") then  ''�ƿ� ���ΰ�� 2016/09/13 �߰�
				    objData = BinaryToText(objXML.ResponseBody, "euc-kr")
			    end if
			end if

			Set objXML  = Nothing

			if (Replace(Trim(objData), vbCrLf, "") <> "") then
				exit for
			end if

			lastipkumdate = searchipkumdate

		next

		if (i >= 20) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[" + CStr(i) + "a]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[1a]"
			response.end
		end if
	else
		'// ========================================================================
		'// �¶��� ���� ���곻��
		response.write "������:::"&CStr(Replace(yyyymmdd, "-", ""))
		xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(Replace(yyyymmdd, "-", "")) + "&key=2beec91670e1f2840a7ac80adde00e49"
		objData = ""
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
		    if (Trim(objXML.ResponseBody)="") then  ''2016/08/22 �߰�
		        response.write "NO_DATA"
		    else
			    objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		    end if
		end if

		Set objXML  = Nothing

		''response.write "aaa" & Trim(objData)

		if (Replace(Trim(objData), vbCrLf, "") = "") then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[--]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[--]"
			response.end
		end if

		searchipkumdate = yyyymmdd
	end if

	''Response.Write objData + "<br>"
	''response.end

	objData = Split(objData, vbCrLf)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

 '' response.write xmlURL
 '' response.end

 	orderserial = requestCheckvar(request("orderserial"),32)
	if (orderserial = "") then
		'// �ߺ� �ֹ���ȣ
		orderserial = "XXXXXXXXX"
	end if

	prevPGkey = ""
	prevPrevPGkey = ""
	prevAppDivCode = ""
	prevPrevAppDivCode = ""
	IsDuplicate = False
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, ";")

		if (UBound(objLine) >= 0) then
			if (objLine(0) = "D") then

				PGgubun		= "uplus"
				PGkey		= objLine(3)
				PGuserid 	= objLine(2)

				if (PGuserid = "dacomtest") then
					sitename = "dacomtest"
				elseif (PGuserid = "tenbyten01") or (PGuserid = "tenbyten02") then
					'// PC MOBILE ���� ����(�ֹ��������� �и�)
					sitename = "10x10"
				else
					sitename = "XXX"
				end if

				if (objLine(6) = "CA01") or (objLine(6) = "CS01") or (objLine(6) = "WR01") then
					'// ==============================
					appDivCode	= "A"
					PGCSkey		= ""

					appDate			= objLine(9)

					cancelDate		= "NULL"
				elseif (objLine(6) = "CA02") or (objLine(6) = "CS02") or (objLine(6) = "WR02") then
					'// ==============================
					appDivCode	= "C"
					PGCSkey		= "CANCELALL"

					appDate			= "NULL"
					cancelDate		= objLine(9)
				elseif (objLine(6) = "CA11") or (objLine(6) = "CS03") or (objLine(6) = "WR06") then
					'// ==============================
					'// �κ����
					'// �������ȯ���� �κ���ҿ� ��ü��Ҹ� ���� �ݾ����� �����ؾ��Ѵ�.
					appDivCode	= "R"
					PGCSkey		= objLine(9) + "-" + objLine(1)			'// �������� + �Ϸù�ȣ

					appDate			= "NULL"
					cancelDate		= objLine(9)
				else
					'// ==============================
					appDivCode = "E"
					PGCSkey		= "ERROR"
				end if

				if (Left(objLine(6), 2) = "CS") then
					appMethod = "7"
				elseif (Left(objLine(6), 2) = "WR") then
					appMethod = "400"
				elseif (Left(objLine(6), 2) = "CA") then
					appMethod = "100"
				else
					appMethod = Left(objLine(6), 2)
				end if

				appPrice		= objLine(7)
				commVatPrice	= round(1.0 * objLine(8) * (1.0/11))
				commPrice		= objLine(8) - commVatPrice
				jungsanPrice	= objLine(7) - objLine(8)

				commPrice = commPrice * -1
				commVatPrice = commVatPrice * -1

				ipkumdate		= objLine(10)

				'// 20130510
				'// (2013-05-10)
				if (appDate <> "NULL") then
					appDate = Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(appDate, 2)
					appDate = "'" + appDate + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(cancelDate, 2)
					cancelDate = "'" + cancelDate + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				''prevPGkey, prevAppDivCode, IsDuplicate

				if (i >= 1) then
					'// �ߺ� ����ó��(���� : 13020397762)
					'// TODO : ������ �ֹ���ȣ ������ ���ĵǾ� �ִٰ� �����Ѵ�.
					'// ���� : 13020397762, 13050293886, 13080752741, 16010731214, 16010731454

					IsDuplicate = False
					If (PGkey = prevPGkey) Then
						if (objLine(6) = "CS01") and (prevAppDivCode = "CS01") Then
							''�ߺ�����
							IsDuplicate = True
						elseif (objLine(6) = "CS02") and (prevAppDivCode = "CS02") Then
							''�ߺ����
							IsDuplicate = True
						elseif (prevPGkey = prevPrevPGkey) Then
							''3���̻�
							IsDuplicate = True
						End If
					End If

					if (prevPGkey <> "") then
						prevPrevPGkey = prevPGkey
						prevPrevAppDivCode = prevAppDivCode
					end if

					prevPGkey = PGkey
					prevAppDivCode = objLine(6)

					if (IsDuplicate = True) Or (PGkey = "18052213890") Or (PGkey = "16010512377") Or (PGkey = "16010731258") Or (PGkey = "18051040230") Or (PGkey = orderserial) then
						sqlStr = " select count(*) as cnt "
						sqlStr = sqlStr + " from "
						sqlStr = sqlStr + " db_temp.dbo.tbl_onlineApp_log_tmp "
						sqlStr = sqlStr + " where "
						sqlStr = sqlStr + " 1 = 1 "
						sqlStr = sqlStr + " and PGkey like '" + CStr(PGkey) + "%' and appDivCode = '" + appDivCode + "' "
						''response.write sqlStr

						subPgKey = ""
						rsget.Open sqlStr,dbget,1
						if Not(rsget.EOF or rsget.BOF) Then
							If rsget("cnt") > 0 Then
								subPgKey = "-" & Format00(2, rsget("cnt"))
							End If
						end if
						rsget.Close

						PGkey = PGkey + subPgKey
					end if
				end if

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				'response.end
				''if (PGkey <> "16010512377") and (PGkey <> "16010512377-01") then
				if PGkey <> "17021377452" then
					dbget.execute sqlStr
				end if
				''end if
			end if
		end if
	Next

	''response.end

	'// ���� : 16010731214
	sqlStr = " update t3 "
	sqlStr = sqlStr + " set t3.PGkey = t1.PGkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t1 "
	sqlStr = sqlStr + " 	left join db_temp.dbo.tbl_onlineApp_log_tmp t2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t1.pggubun = t2.pggubun "
	sqlStr = sqlStr + " 		and t1.PGkey = t2.PGkey "
	sqlStr = sqlStr + " 		and t2.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp t3 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t1.pggubun = t3.pggubun "
	sqlStr = sqlStr + " 		and Left(t1.PGkey, 11) = t3.PGkey "
	sqlStr = sqlStr + " 		and t3.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t1.PGgubun = 'uplus' "
	sqlStr = sqlStr + " 	and Len(t1.PGkey) > 11 "
	sqlStr = sqlStr + " 	and t1.PGCSkey = '' "
	sqlStr = sqlStr + " 	and t2.PGkey is NULL "
	dbget.execute sqlStr

	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set orderserial = pgkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(pgkey) < 20 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ӽ��ֹ���ȣ => �ֹ���ȣ
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex('|', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex('|', o.paygatetid) > 0 "										'// ������ '|'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ӽ��ֹ���ȣ => �ֹ���ȣ
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex(';', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex(';', o.paygatetid) > 0 "										'// ������ ';'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �����
	sqlStr = " update t set t.sitename = '10x10mobile' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	''sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	''sqlStr = sqlStr + " 			and o.rdsite = 'mobile' "													'// �����
	sqlStr = sqlStr + " 			and o.beadaldiv in (4,5,7,8) "												'// �����(4:mobile, 5:mobile_link, 7:APP, 8:between ), 2015-07-13, skyer9
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// PG�� ������
	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set PGmeachulDate = convert(varchar(10), IsNull(cancelDate, appdate), 127) "
	sqlStr = sqlStr + " where pggubun = 'uplus' and PGmeachulDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ǰ�����
	sqlStr = " update t set t.appdate = IsNull(o.ipkumdate, t.appdate) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where t.pggubun = 'uplus' and appDivCode = 'A' and o.jumundiv not in ('6', '9') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �������
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "						'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �������(��ȯ�ֹ� ��ǰ)
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = c.orgorderserial "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''sqlStr = " delete from db_order.dbo.tbl_onlineApp_log where PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	sqlStr = " delete l "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " 		and t.PGgubun = l.PGgubun "
	sqlStr = sqlStr + " 		and t.PGkey = l.PGkey "
	sqlStr = sqlStr + " 		and t.PGCSkey = l.PGCSkey "
	sqlStr = sqlStr + " 		and l.PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��Ī
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.pgkey = o.orderserial "
	sqlStr = sqlStr + " where l.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('�ŷ����� : " + CStr(searchipkumdate) + "');</script>"
	end If

elseif (mode = "matchpgdata") then

	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco') "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	''sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX')  or (l.PGgubun = 'payco')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'' 6���� ���� �������� ��Ī
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_log.dbo.tbl_old_order_master_2003 o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "

	'/ã�� ���� PGkey �̰͵� ������ �־����	'/2016.09.22 �ѿ��
	sqlStr = sqlStr + " 	and l.PGkey in ('IniTechPG_teenxteen420170516133329200702', 'INIMX_CARDteenxteen920170118085058051814', 'INIMX_CARDteenxteen920170410101203081721') "

	'' �ϴ� ���� �ʿ��� �� ����� ��������(2014-09-05, skyer9)
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	'// ���ֹ� �������
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "								'// ��Ҹ�
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// ��Ҵ� �Ѱ��̹Ƿ� �Ϸ� ���̳��� ��Ī
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay', 'naverpay', 'payco') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� �������(OK+�ſ�)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// ��Ҵ� �Ѱ��̹Ƿ� �Ϸ� ���̳��� ��Ī
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��

	'/2���� ������� �ؿ��� 4��¥�� �ּ� ���� Ǯ�� �ְ� ������.
	'sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) < 4 "		'// 4��
	'sqlStr = sqlStr + " 	and l.canceldate >= '2016-12-01' "		'/��¥�� �ٲ��ְ�

	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay', 'naverpay', 'payco') "

	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ(OK+�ſ�)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��ȯ�ֹ� ��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.orderserial = c.orgorderserial and c.deldate is NULL "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''�ߺ���Ī Ȯ��
	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'A'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'C'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, csasid, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode <> 'A' and csasid is not NULL
	'' group by orderserial, csasid
	'' having count(*) > 1


	'// �κ�����̸鼭 �������� ����� ���
	'// cancelDate �� ������ ���� ���ڷ� �����ǰ� �ð��븸 �����ϰ� �����ȴ�.
	'// ���� �ð��븸 ���ؼ� ��Ī���ش�.
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	left join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate) % (24 * 60)) < 2 "			'// ���� �ð���
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''dbget.close
	''Response.End

elseif (mode = "matchpgdata6month") then
	'// 6�������� ���� ��Ī(PG Key �ִ� ��츸)

	PGkey = requestCheckvar(request("PGkey"),64)

	'' 6���� ���� �������� ��Ī
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_log.dbo.tbl_old_order_master_2003 o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
	sqlStr = sqlStr + " 	and l.PGkey = '" & PGkey & "' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ֹ����
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGkey = '" & PGkey & "' "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �κ����
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	left join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGkey = '" & PGkey & "' "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate) % (24 * 60)) < 2 "			'// ���� �ð���
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchfingerspgdata") then

	'' sqlStr = " update l set l.orderserial = o.orderserial "
	'' sqlStr = sqlStr + " from "
	'' sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master o "
	'' sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	'' sqlStr = sqlStr + " on "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	'' sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	'' sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	'' sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	'' sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	'' sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	'' sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX')) "
	'' sqlStr = sqlStr + " where "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and l.orderserial is NULL "
	sqlStr = " exec [db_order].[dbo].[usp_TEN_PGData_Match_FingersOrder] "
	''response.write sqlStr + "<br>"
	''response.end
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].dbo.tbl_academy_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��

	'/2���� ������� �ؿ��� 4��¥�� �ּ� ���� Ǯ�� �ְ� ������.
	'sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 4 "		'// 4��
	'sqlStr = sqlStr + " 	and l.canceldate >= '2016-12-01' "		'/��¥�� �ٲ��ְ�

	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'kcp') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchgiftcardpgdata") then
    ''�̴Ͻý� ��������� ��� �Աݿ�û TID��  �ԱݿϷ�TID�� �ٸ�����  tbl_onlineApp_log ���� �Ա�TID�� ��
	sqlStr = " update l set l.orderserial = o.giftorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_giftcard_order o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paydateid = l.PGkey "
	sqlStr = sqlStr + " 	and o.paydateid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) <= 5) or (l.appDivCode <> 'A'))  "		'// 5��
	sqlStr = sqlStr + " 	and ((Left(o.paydateid, 9) = 'IniTechPG') or (Left(o.paydateid, 5) = 'INIMX') or (Left(o.paydateid, 6) = 'INIswt') or (Left(o.paydateid, 6) = 'Stdpay') or (Left(o.paydateid, 5) = 'StdMX')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and IsNull(l.orderserial, '') = '' "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �Աݿ�û TID �� �����Ա�TID �� ���� �ٸ���.
	'// db_order.dbo.tbl_cyberAcctNoti_LogINI �� �̿��ؼ� ��Ī�����ش�.
	sqlStr = " update l "
	sqlStr = sqlStr + " set l.orderserial = o.giftorderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_giftcard_order o "
	sqlStr = sqlStr + " 	Join db_order.dbo.tbl_cyberAcctNoti_LogINI Nt "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and Nt.NO_OID=o.giftorderserial "
	sqlStr = sqlStr + " 		and Nt.isMatched='Y' "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and Nt.NO_TID = l.PGkey "
	sqlStr = sqlStr + " 		and o.paydateid is not NULL "
	sqlStr = sqlStr + " 		and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 		and o.jumundiv <> '6' "
	sqlStr = sqlStr + " 		and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) <= 5) or (l.appDivCode <> 'A')) "
	sqlStr = sqlStr + " 		and ((Left(l.PGkey , 9) = 'IniTechPG') or (Left(l.PGkey, 5) = 'INIMX') or (Left(l.PGkey, 6) = 'INIswt') or (Left(l.PGkey, 6) = 'Stdpay') or (Left(l.PGkey, 5) = 'StdMX')) "
	sqlStr = sqlStr + " 	and appmethod='7' "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and IsNull(l.orderserial, '') = '' "
	sqlStr = sqlStr + " 		and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) <= 5 "		'// 5��
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��ü���� �� ������ ������Ʈ
	sqlStr = " update "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orgPayDate = convert(VARCHAR(10), appDate, 127) "
	sqlStr = sqlStr + " where appDate is not NULL and orgPayDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.orgPayDate = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log r "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and r.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	and r.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	and r.appDivCode = 'R' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where r.appDate is NULL and a.appDate is not NULL and r.orgPayDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "makeadvprc") then

	sqlStr = " select PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[DBDATAMART].db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and paydate is not NULL "
	sqlStr = sqlStr + " 	and pgkey is not NULL "
	sqlStr = sqlStr + " 	and paydate >= '" + Left(DateAdd("m", -1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and paydate < '" + Left(DateAdd("m", 1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and payDivCode not in ('mil', 'dep', 'gif', '0', 'XXX') "
	sqlStr = sqlStr + " 	and not (payDivCode in ('rde') and realPayPrice = 0) "
	sqlStr = sqlStr + " 	and PGgubun <> 'KICC' "

	'// ���� ��¥ ���� �ݾ��� ȯ�Ұ��� �ִ� ��� �߸� ��Ī�� �� �ִ�.
	'// sqlStr = sqlStr + " 	and PGkey<>'15062692753'"  ''�ϴ�.����. �� �ذ�.

	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " having count(*) > 1 "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		response.write "����α� �ߺ� ERROR : PGKey = " & rsget("pgkey")
		rsget.close
		dbget.close()
		response.end
	end if
	rsget.close

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake] '" + CStr(yyyymm) + "' "
	''rw sqlStr : response.end
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('�ۼ��Ǿ����ϴ�.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

elseif (mode = "makeadvprc01") then

	sqlStr = " select PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[DBDATAMART].db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and paydate is not NULL "
	sqlStr = sqlStr + " 	and pgkey is not NULL "
	sqlStr = sqlStr + " 	and paydate >= '" + Left(DateAdd("m", -1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and paydate < '" + Left(DateAdd("m", 1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and payDivCode not in ('mil', 'dep', 'gif', '0', 'XXX') "
	sqlStr = sqlStr + " 	and not (payDivCode in ('rde') and realPayPrice = 0) "
	sqlStr = sqlStr + " 	and PGgubun <> 'KICC' "

	'// ���� ��¥ ���� �ݾ��� ȯ�Ұ��� �ִ� ��� �߸� ��Ī�� �� �ִ�.
	'// sqlStr = sqlStr + " 	and PGkey<>'15062692753'"  ''�ϴ�.����. �� �ذ�.

	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " having count(*) > 1 "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		response.write "����α� �ߺ� ERROR : PGKey = " & rsget("pgkey")
		rsget.close
		dbget.close()
		response.end
	end if
	rsget.close

elseif (mode = "makeadvprc02") then

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake] '" + CStr(yyyymm) + "' "
	''rw sqlStr : response.end
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('�ۼ��Ǿ����ϴ�.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

end if

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script language='javascript'>
alert('����Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->