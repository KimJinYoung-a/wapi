<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Server.ScriptTimeOut = 60*60 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->

<!-- #include virtual="/lib/function.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'#######################################################
'	History	: 2017.11.02 원승현 생성
'	Description : appBoy에 유저정보 전송
'#######################################################
	Dim mode, userId, oXML, sendJsonData, query1, appBoySendEnable, myJson, query2, platForm, itemids, limitUseq, totalcnt, i
	'// 앱보이에 넘길값들의 변수
	Dim username, usermail, usermileage, dob, gender, firstLoginDate, lastLoginDate, userLevel, external_id, push_subscribe, push_opted_in_at, appBoyItemName
	'// 품절상품 입고알림에 사용될 상품변수
	Dim productName, itemId, productImage, iOsMessage, aOsMessage


	mode = request("mode")
	userId = request("userid")
	platForm = request("platForm")
	itemids = request("itemids")

	appBoySendEnable = False


	If Trim(mode)="" Then
		response.End
	End If

	Select Case Trim(mode)

		'// 매일 아침 6시 사용자 일반정보(생일, 성별, 첫번째로그인, 마지막로그인, 등급, 푸쉬 수신여부, 푸쉬 수신허용일자, 유저이메일, 유저이름)
		Case "usergeneralinfo"
			'// 업데이트할 전체 인원과 현재까지 등록된 useq값을 가져온다.
			query1 = query1 + "	Select count(n.userid) as totalcnt, max(useq) as useq"
			query1 = query1 + "	From db_user.dbo.tbl_user_n n"
			query1 = query1 + "	inner join db_user.dbo.tbl_logindata l on n.userid = l.userid"
			query1 = query1 + "	left join db_contents.dbo.tbl_app_wish_userinfo u on n.userid = u.userid"
			'dbDatamart_rsget.CommandTimeOut = 480
			dbDatamart_rsget.Open query1,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
			If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
				totalcnt = dbDatamart_rsget("totalcnt")
				limitUseq =  dbDatamart_rsget("useq")
			End If
			dbDatamart_rsget.close

			For i = 1 To totalcnt Step 50
				query2 = " Select "
				query2 = query2 + " 	num, userid, dob, usermail, firstLogin, lastLogin, username, gender, userlevel, external_id, push_subscribe, push_opted_in_at "
				query2 = query2 + " From "
				query2 = query2 + " ( "
				query2 = query2 + " 	Select  "
				query2 = query2 + " 		ROW_NUMBER() over(order by n.userid) as num, "
				query2 = query2 + " 		n.userid, n.username, "
				query2 = query2 + " 		case when convert(varchar(10), birthday, 120)='1900-01-01' then 'null' else convert(varchar(10), birthday, 120) end as dob, "
				query2 = query2 + " 		usermail, convert(varchar(33), regdate, 126)+'+09:00' as firstLogin, convert(varchar(33), l.lastlogin, 126)+'+09:00' as lastLogin,  "
				query2 = query2 + " 		case when n.sexflag in (1,3,5,7) then 'M' when n.sexflag in (2,4,6,8) then 'F' else 'null' end as gender,  "
				query2 = query2 + " 		case when userlevel=0 then 'yellow' when userlevel=1 then 'green' when userlevel=2 then 'blue' when userlevel=3 then 'vipsilver' "
				query2 = query2 + " 			when userlevel=4 then 'vipgold' when userlevel=5 then 'orange' when userlevel=6 then 'vvip' when userlevel=7 then 'staff'  "
				query2 = query2 + " 			when userlevel=8 then 'family' end as userlevel, "
				query2 = query2 + " 		useq*3 as external_id,  "
				query2 = query2 + " 		case when lastpushyn='Y' then 'opted_in' when lastpushyn='N' then 'unsubscribed' else 'subscribed' end as push_subscribe,  "
				query2 = query2 + " 		case when lastpushyn='Y' then convert(varchar(33), lastpushynDate, 126)+'+09:00' else 'null' end as push_opted_in_at "
				query2 = query2 + " 	From db_user.dbo.tbl_user_n n "
				query2 = query2 + " 	inner join db_user.dbo.tbl_logindata l on n.userid = l.userid "
				query2 = query2 + " 	left join db_contents.dbo.tbl_app_wish_userinfo u on n.userid = u.userid "
				query2 = query2 + " 	Where l.useq <= "&limitUseq
				query2 = query2 + " )AA "
				query2 = query2 + " Where num >= "&i&" And num < "&i+50
				'dbget.CommandTimeOut = 480
				dbDatamart_rsget.Open query2,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
				If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
					appBoySendEnable = True
					sendJsonData = ""
					Do Until dbDatamart_rsget.eof
						'// appboy에 넘길값을 json형태로 만듬.
						sendJsonData = sendJsonData & "{"
						sendJsonData = sendJsonData & """dob"":"""&dbDatamart_rsget("dob")&"""" '// 생년월일
						sendJsonData = sendJsonData & ",""gender"":"""&dbDatamart_rsget("gender")&"""" '// 성별
						sendJsonData = sendJsonData & ",""firstLoginDate"":"""&dbDatamart_rsget("firstLogin")&"""" '// 첫번째 로그인
						sendJsonData = sendJsonData & ",""lastLoginDate"":"""&dbDatamart_rsget("lastLogin")&"""" '// 마지막 로그인
						sendJsonData = sendJsonData & ",""userlevel"":"""&dbDatamart_rsget("userlevel")&"""" '// 회원등급
						sendJsonData = sendJsonData & ",""external_id"":"""&dbDatamart_rsget("external_id")&"""" '// appboy용 회원아이디
						sendJsonData = sendJsonData & ",""push_subscribe"":"""&dbDatamart_rsget("push_subscribe")&"""" '// 앱Push 허용여부
						sendJsonData = sendJsonData & ",""push_opted_in_at"":"""&dbDatamart_rsget("push_opted_in_at")&"""" '// 앱Push 허용일자
						sendJsonData = sendJsonData & ",""email"":"""&dbDatamart_rsget("usermail")&"""" '// 유저이메일
						sendJsonData = sendJsonData & ",""username"":"""&dbDatamart_rsget("username")&"""" '// 유저이름
						sendJsonData = sendJsonData & "},"
					dbDatamart_rsget.movenext
					Loop
				End If
				dbDatamart_rsget.close

				If appBoySendEnable Then
					sendJsonData = Left(sendJsonData, Len(sendJsonData)-1)


					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.appboy.com/users/track", true
					oXML.setRequestHeader "Content-Type", "application/json; charset=utf-8"
					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
					oXML.setRequestHeader "Accept","application/json"
					oXML.send "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""attributes"":["&sendJsonData&"]}"	'바디 전송

					'// 여기까지만 처리하고 상단 부분을 비동기로 바꿔보자.. 속도가 너무 안나오는데..
		
					'response.write oXML.responseText
					'response.End
					'Set myJson = JSON.parse(oXML.responseText)
					'response.write myJson.message
				
					'If Trim(myJson.message) <> "success" Then
					'	response.write "오류가 발생했습니다.<br>"&myJson.message
					'	Set oXML = Nothing
					'	Set myJson = Nothing
					'	appBoySendEnable = False
					'	Exit For
					'End If
					response.write i
					appBoySendEnable = False

					Set oXML = Nothing
					'Set myJson = Nothing
				End If

			Next

			'// 성공 실패 여부를 db에 담는다.
			query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', '"&myJson.message&"', '"&i&"',  getdate())"
			dbCTget.execute query2
			response.write "AppBoy로 유저 기본정보 전송이 완료되었습니다."

		'// 매일오후 12시 50분 toDay 할인시작된 상품을 가진 유저정보
		Case "UserToDaySaleBasket"
			query1 = query1 + "	Select count(*) as totalcnt"
			query1 = query1 + "	From db_dumi.dbo.tbl_appBoyBasketData "
			'dbDatamart_rsget.CommandTimeOut = 480
			dbDatamart_rsget.Open query1,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
			If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
				totalcnt = dbDatamart_rsget("totalcnt")
			End If
			dbDatamart_rsget.close

			For i = 1 To totalcnt Step 50

				query2 = " Select "
				query2 = query2 + " num, userid, external_id*3 as external_id, itemid, itemname, convert(varchar(33), regdate, 126)+'+09:00' as saleItemInBasketDate "
				query2 = query2 + " From db_dumi.dbo.tbl_appBoyBasketData "
				query2 = query2 + " Where num >= "&i&" And num < "&i+50
				dbDatamart_rsget.Open query2,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
				If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
					appBoySendEnable = True
					sendJsonData = ""
					Do Until dbDatamart_rsget.eof
						If Len(dbDatamart_rsget("itemname"))>=12 Then
							appBoyItemName = Left(dbDatamart_rsget("itemname"), 12)&".."
						Else
							appBoyItemName = dbDatamart_rsget("itemname")
						End If
						'// appboy에 넘길값을 json형태로 만듬.
						sendJsonData = sendJsonData & "{"
						sendJsonData = sendJsonData & """external_id"":"""&dbDatamart_rsget("external_id")&"""" '// appboy용 회원아이디
						sendJsonData = sendJsonData & ",""saleItemInBasketDate"":"""&dbDatamart_rsget("saleItemInBasketDate")&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(오늘날짜로 셋팅하여 푸시 보내면됨)
						sendJsonData = sendJsonData & ",""saleItemInBasketItemId"":"""&dbDatamart_rsget("itemid")&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(상품코드)
						sendJsonData = sendJsonData & ",""saleItemInBasketItemName"":"""&appBoyItemName&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(상품명)
						sendJsonData = sendJsonData & "},"
					dbDatamart_rsget.movenext
					Loop
				End If
				dbDatamart_rsget.close

				If appBoySendEnable Then
					sendJsonData = Left(sendJsonData, Len(sendJsonData)-1)

					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.appboy.com/users/track", False
					oXML.setRequestHeader "Content-Type", "application/json; charset=utf-8"
					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
					oXML.setRequestHeader "Accept","application/json"
					oXML.send "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""attributes"":["&sendJsonData&"]}"	'바디 전송

					'Set myJson = JSON.parse(oXML.responseText)

					If InStr(oXML.responseText, "success") < 1 Then
						'// 성공 실패 여부를 db에 담는다.
						query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'failed', '"&i&"',  getdate())"
						dbCTget.execute query2					
						Set oXML = Nothing
						Set myJson = Nothing
						appBoySendEnable = False
						Exit For
					End If

					appBoySendEnable = False
				End If
			Next
			'// 성공 실패 여부를 db에 담는다.
			query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'success', '"&i&"',  getdate())"
			dbCTget.execute query2
			response.write "AppBoy로 장바구니에 있는 세일상품정보가 전송이 완료되었습니다."

			Set oXML = Nothing
			'Set myJson = Nothing

		'// 매일오후 1시 10분 toDay 할인시작된 상품을 가진 유저정보
		Case "UserToDaySaleWish"
			query1 = query1 + "	Select count(*) as totalcnt"
			query1 = query1 + "	From db_dumi.dbo.tbl_appBoyWishData "
			'dbDatamart_rsget.CommandTimeOut = 480
			dbDatamart_rsget.Open query1,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
			If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
				totalcnt = dbDatamart_rsget("totalcnt")
			End If
			dbDatamart_rsget.close

			For i = 1 To totalcnt Step 50

				query2 = " Select "
				query2 = query2 + " num, userid, external_id*3 as external_id, itemid, itemname, convert(varchar(33), regdate, 126)+'+09:00' as saleItemInWishDate "
				query2 = query2 + " From db_dumi.dbo.tbl_appBoyWishData "
				query2 = query2 + " Where num >= "&i&" And num < "&i+50
				dbDatamart_rsget.Open query2,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
				If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
					appBoySendEnable = True
					sendJsonData = ""
					Do Until dbDatamart_rsget.eof
						If Len(dbDatamart_rsget("itemname"))>=12 Then
							appBoyItemName = Left(dbDatamart_rsget("itemname"), 12)&".."
						Else
							appBoyItemName = dbDatamart_rsget("itemname")
						End If
						'// appboy에 넘길값을 json형태로 만듬.
						sendJsonData = sendJsonData & "{"
						sendJsonData = sendJsonData & """external_id"":"""&dbDatamart_rsget("external_id")&"""" '// appboy용 회원아이디
						sendJsonData = sendJsonData & ",""saleItemInWishDate"":"""&dbDatamart_rsget("saleItemInWishDate")&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(오늘날짜로 셋팅하여 푸시 보내면됨)
						sendJsonData = sendJsonData & ",""saleItemInWishItemId"":"""&dbDatamart_rsget("itemid")&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(상품코드)
						sendJsonData = sendJsonData & ",""saleItemInWishItemName"":"""&appBoyItemName&"""" '// 할인상품을 가지고 있는 유저만 들어가는값(상품명)
						sendJsonData = sendJsonData & "},"
					dbDatamart_rsget.movenext
					Loop
				End If
				dbDatamart_rsget.close

				If appBoySendEnable Then
					sendJsonData = Left(sendJsonData, Len(sendJsonData)-1)

					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.appboy.com/users/track", False
					oXML.setRequestHeader "Content-Type", "application/json; charset=utf-8"
					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
					oXML.setRequestHeader "Accept","application/json"
					oXML.send "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""attributes"":["&sendJsonData&"]}"	'바디 전송

					'Set myJson = JSON.parse(oXML.responseText)

					If InStr(oXML.responseText, "success") < 1 Then
						'// 성공 실패 여부를 db에 담는다.
						query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'failed', '"&i&"',  getdate())"
						dbCTget.execute query2					
						Set oXML = Nothing
						Set myJson = Nothing
						appBoySendEnable = False
						Exit For
					End If

					appBoySendEnable = False
				End If
			Next
			'// 성공 실패 여부를 db에 담는다.
			query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'success', '"&i&"',  getdate())"
			dbCTget.execute query2
			response.write "AppBoy로 Wish에 있는 세일상품정보가 전송이 완료되었습니다."

			Set oXML = Nothing
			'Set myJson = Nothing

		'// 매일새벽 12시 30분 탈퇴회원을 appboy에서 삭제해준다.
		Case "DelUser"
			query1 = query1 + "	Select count(userid) as totalcnt"
			query1 = query1 + "	From db_user.[dbo].[tbl_deluser] "
			query1 = query1 + "	Where useq is not null And convert(varchar(10), regdate, 120) = convert(varchar(10), dateadd(day, -1, getdate()), 120) "
			'dbDatamart_rsget.CommandTimeOut = 480
			rsget.Open query1,dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.bof Or rsget.eof) Then
				totalcnt = rsget("totalcnt")
			End If
			rsget.close

			For i = 1 To totalcnt Step 50
				query2 = " Select num, external_id "
				query2 = query2 + " From "
				query2 = query2 + " ( "
				query2 = query2 + " 	Select ROW_NUMBER() over(order by useq) as num, useq*3 as external_id "
				query2 = query2 + " 	From db_user.[dbo].[tbl_deluser] "
				query2 = query2 + " 	Where useq is not null And convert(varchar(10), regdate, 120) = convert(varchar(10), dateadd(day, -1, getdate()), 120) "
				query2 = query2 + " )AA Where num >= "&i&" And num < "&i+50
				rsget.Open query2,dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.bof Or rsget.eof) Then
					appBoySendEnable = True
					sendJsonData = ""
					Do Until rsget.eof
						'// appboy에 넘길값을 json형태로 만듬.
						sendJsonData = sendJsonData & """"&rsget("external_id")&"""," '// appboy용 회원아이디
					rsget.movenext
					Loop
				End If
				rsget.close

				If appBoySendEnable Then
					sendJsonData = Left(sendJsonData, Len(sendJsonData)-1)

					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.appboy.com/users/delete", False
					oXML.setRequestHeader "Content-Type", "application/json; charset=utf-8"
					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
					oXML.setRequestHeader "Accept","application/json"
					oXML.send "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""external_ids"":["&sendJsonData&"]}"	'바디 전송

					'Set myJson = JSON.parse(oXML.responseText)

					If InStr(oXML.responseText, "success") < 1 Then
						'// 성공 실패 여부를 db에 담는다.
						query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'failed', '"&i&"',  getdate())"
						dbCTget.execute query2					
						Set oXML = Nothing
						Set myJson = Nothing
						appBoySendEnable = False
						Exit For
					End If

					appBoySendEnable = False
					Set oXML = Nothing
				End If
			Next
			'// 성공 실패 여부를 db에 담는다.
			query2 = " insert into db_apiLog.dbo.tbl_AppBoyApiSendLog (mode, returnMessage, numSideSend, regdate) values ('"&mode&"', 'success', '"&i&"',  getdate())"
			dbCTget.execute query2
			response.write "텐바이텐 탈퇴회원을 AppBoy에서 삭제하였습니다."
			response.End

		Case "test"
		
			'// 업데이트할 전체 인원과 현재까지 등록된 useq값을 가져온다.
			query1 = query1 + "	Select count(n.userid) as totalcnt, max(useq) as useq"
			query1 = query1 + "	From db_user.dbo.tbl_user_n n"
			query1 = query1 + "	inner join db_user.dbo.tbl_logindata l on n.userid = l.userid"
			query1 = query1 + "	left join db_contents.dbo.tbl_app_wish_userinfo u on n.userid = u.userid"
			dbDatamart_rsget.CommandTimeOut = 480
			dbDatamart_rsget.Open query1,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
			If Not(dbDatamart_rsget.bof Or dbDatamart_rsget.eof) Then
				totalcnt = dbDatamart_rsget("totalcnt")
				limitUseq =  dbDatamart_rsget("useq")
			End If
			dbDatamart_rsget.close

			For i = 1 To totalcnt Step 50

				response.write r&":"&i&"-"&i+50&"<br>"

				r = r+1

			Next

		'// 품절상품 입고알림 푸시. 루틴하게 오전9시, 오후 12시, 오후 6시에 발송한다.
		Case "soldoutpush"
	
			query1 = ""
			query1 = query1 + "	Select SA.Idx, SA.itemid, SA.ItemOptionCode, SA.UserId, SA.External_Id, SA.AlamType, SA.Regdate, SA.LimitPushDate "
			query1 = query1 + "		, i.itemname, "
			query1 = query1 + "		'http://thumbnail.10x10.co.kr/webimage/image/list/'+ "
			query1 = query1 + "			CASE WHEN LEN(CONVERT(VARCHAR(20),(SA.itemid / 10000)))=1 THEN '0'+convert(VARCHAR(20),(SA.itemid / 10000)) ELSE CONVERT(VARCHAR(20),(SA.itemid / 10000)) END+ "
			query1 = query1 + "			'/'+i.listimage AS listimage	 "
			query1 = query1 + "		, o.optionname "
			query1 = query1 + "	From db_my10x10.[dbo].[tbl_SoldOutProductAlam] SA with (nolock) "
			query1 = query1 + "	inner join db_item.dbo.tbl_item i with (nolock) on SA.itemid = i.itemid And i.sellyn='Y' "
			query1 = query1 + "	left join db_item.dbo.tbl_item_option o with (nolock) on SA.itemid = o.itemid And SA.ItemOptionCode = o.itemoption And o.isusing='Y' And o.optsellyn='Y' And o.optlimityn='Y' "
			query1 = query1 + "	Where SA.Idx is not null And "
			query1 = query1 + "		case when i.optioncnt=0 then (i.limitno - i.limitsold) else (o.optlimitno - o.optlimitsold) end > 0 "
			query1 = query1 + "		And getdate() < LimitPushDate And SA.SendPushDate is null And SA.SendStatus='N' "
			rsget.Open query1,dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.bof Or rsget.eof) Then
				Do Until rsget.eof
					'// DB에서 사용자 external_ids값, 품절이 풀린 상품명, 상품이미지, 상품코드를 가져온다.
					external_id = """"&rsget("External_Id")&""""
					productName = rsget("itemname")
					itemId = rsget("ItemId")
					productImage = rsget("listimage")
					iOsMessage = "{"
					iOsMessage = iOsMessage &"""alert"":""고객님께서 입고 알림 신청하신 "&productName&"가 구매가능합니다. 구매시점에 따라 품절이 될 수 있으니 서둘러주세요!"""
					iOsMessage = iOsMessage &",""custom_uri"":""tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&itemId&"&rdsite=appBoy"""
					iOsMessage = iOsMessage &",""asset_url"":"""&productImage&""""
					iOsMessage = iOsMessage &",""asset_file_type"":""jpg"""
					iOsMessage = iOsMessage &"}"
					aOsMessage = "{"
					aOsMessage = aOsMessage &"""alert"":""고객님께서 입고 알림 신청하신 "&productName&"가 구매가능합니다. 구매시점에 따라 품절이 될 수 있으니 서둘러주세요!"""
					aOsMessage = aOsMessage &",""custom_uri"":""tenwishapp://http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid="&itemId&"&rdsite=appBoy"""
					aOsMessage = aOsMessage &",""push_icon_image_url"":"""&productImage&""""
					aOsMessage = aOsMessage &"}"

					set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
					oXML.open "POST", "https://api.appboy.com/messages/send", False
					oXML.setRequestHeader "Content-Type", "application/json; charset=utf-8"
					oXML.setRequestHeader "CharSet", "utf-8" '있어도 되고 없어도 되고
					oXML.setRequestHeader "Accept","application/json"
					oXML.send "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""broadcast"":false,""override_frequency_capping"":true,""external_user_ids"":["&external_id&"], ""messages"":{""apple_push"":"&iOsMessage&", ""android_push"":"&aOsMessage&"}}"

					If InStr(oXML.responseText, "success") > 0 Then
						'// 여기서 품절입고알림 신청 테이블 정보를 업데이트 해준다.
						query2 = " update db_my10x10.[dbo].[tbl_SoldOutProductAlam] set SendPushDate = getdate(), SendStatus='Y' Where idx = '"&rsget("Idx")&"' "
						dbget.execute query2
					End If

					Set oXML = Nothing
				Loop
				rsget.close
			End If

'			response.write external_id&"<br>"&productName&"<br>"&itemId&"<br>"&productImage&"<br>"&iOsMessage
'			response.write "{""app_group_id"":""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""broadcast"":false,""override_frequency_capping"":true,""external_user_ids"":["&external_id&"], ""messages"":{""apple_push"":"&iOsMessage&"}}"
'			response.End

			'oXML.send "{""app_group_id"": ""9dca85e3-8cf2-406a-a98b-d2cba7d2d3df"",""broadcast"" : false, ""override_frequency_capping"" : true, ""external_user_ids"" : [""38706282""], messages: { apple_push: { alert: ""dd""}}}"

			response.write "Push Message가 발송되었습니다."
			response.End

	End Select



	Function Ceil(ByVal intParam)
	 Ceil = -(Int(-(intParam)))
	End Function
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbDatamartclose.asp" -->