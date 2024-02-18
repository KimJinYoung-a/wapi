<%

'// ============================================================================
'// 77번 디비 필요
'// /lib/db/dbAppNotiopen.asp
'// /lib/db/dbAppNoticlose.asp
'// ============================================================================

'' msggroup_id			'// 그룹메시지

'' sender_regid			'// 예비용(사용안함)
'' sender_userid
'' sender_role

'' receiver_regid		'// 수신자 등록코드
'' receiver_userid		'// 아이디
'' receiver_role		'// 권한

'' message				'// 메시지

'' url					'// 전달할 URL

'' 권한
'' Const ROLE_UNKNOWN = -1
'' Const ROLE_MEMBER = 10
'' Const ROLE_STAFF = 20

'// 푸시알림 메세지를 등록하는 함수
Function sendPushMessage(appkey, receiver_userid, receiver_role, message, url)
	dim sqlStr, AssignedRow, tmpRS
	dim msggroup_id

	'// 사용법
	'// 1. 특정인
	'// 	Call sendPushMessage("admin_app", "skyer9", "20", "hello, world", "http://www.10x10.co.kr/")
	'// 2. 전체
	'// 	Call sendPushMessage("admin_app", "*", "20", "hello, world", "http://www.10x10.co.kr/")

	sendPushMessage = False

	msggroup_id = -1

	appkey = Replace(appkey, "'", "")
	receiver_userid = Replace(receiver_userid, "'", "")
	receiver_role = Replace(receiver_role, "'", "")
	message = Replace(message, "'", "")

	appkey = Replace(appkey, "--", "")
	receiver_userid = Replace(receiver_userid, "--", "")
	receiver_role = Replace(receiver_role, "--", "")
	message = Replace(message, "--", "")

	if (receiver_userid = "*") then
		sqlStr = "INSERT INTO db_AppNoti.dbo.tbl_tbtpns_msggroup (complete, message, url) VALUES (0, '" & message & "' , '" & url & "') "
		''response.write sqlStr
		dbAppNotiget.Execute(sqlStr)

		'// TODO : 프로시져로 변경 필요
		sqlStr = " SELECT @@IDENTITY AS NewID "
		Set tmpRS = dbAppNotiget.Execute(sqlStr)
		msggroup_id = tmpRS.Fields("NewID").value
		tmpRS.close
		Set tmpRS = Nothing
	end if

	sqlStr = " insert into db_AppNoti.dbo.tbl_tbtpns_msgitem(msggroup_id, sender_regid, sender_userid, sender_role, receiver_regid, receiver_userid, receiver_role, message, url, resultcode, resultmessage, resultdate) "
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " 	" + CStr(msggroup_id) + " "
	sqlStr = sqlStr + " 	, -1, '', -1 "
	sqlStr = sqlStr + " 	, r.regid, r.userid, r.role "
	sqlStr = sqlStr + " 	, '" + CStr(message) + "' "
	sqlStr = sqlStr + " 	, '" + CStr(url) + "' "
	sqlStr = sqlStr + " 	, 100, '진행', getdate() "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_AppNoti.dbo.tbl_tbtpns_app a "
	sqlStr = sqlStr + " 	join db_AppNoti.dbo.tbl_tbtpns_register r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.appid = r.appid "
	sqlStr = sqlStr + " 		and a.allowed_role = r.role "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.useyn = 'Y' "
	sqlStr = sqlStr + " 	and a.appkey = '" + CStr(appkey) + "' "

	if (receiver_userid <> "*") then
		sqlStr = sqlStr + " 	and r.userid = '" + CStr(receiver_userid) + "' "
	end if

	sqlStr = sqlStr + " 	and r.role = '" + CStr(receiver_role) + "' "
	sqlStr = sqlStr + " order by r.regid "
	dbAppNotiget.Execute sqlStr, AssignedRow

	sendPushMessage = (AssignedRow > 0)
End Function

%>
