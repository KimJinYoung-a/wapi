<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/twitter/ASPTwitter/ASPTwitter.asp"-->
<!-- #include virtual="/lib/util/dbCacheLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'#######################################################
'	History	: 2015.01.22 원승현 생성
'	Description : 텐바이텐 sns 가져오기
'#######################################################

dim refIP : refIP = request.serverVariables("REMOTE_ADDR")


if NOT ((LEFT(refIP,10)="61.252.133") or (refIP="110.93.128.113") or (refIP="110.93.128.113")) then
    response.write "ERR-01"
end if
 

On Error Resume Next

Dim fbSns, instaSns, twSns, twitter_api_consumer_key, twitter_api_consumer_secret, objASPTwitter, oTweet, fnSnsImg
Dim  i, j
	j = 0

dim iTxtVal

	iTxtVal = getSnsTxtValue()
	
	Call SetDBCacheTxtVal("wSnsCache", iTxtVal, "", "21600")
	
	if Err then
        response.write Err.description	    
	ELSE
	    response.write "OK:"&LEN(iTxtVal)
    END IF

On Error goto 0

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	'// asp json 거시기
	Function getJsonAsp(url, param)
		Dim objHttp
		Dim strJsonText
		Set objHttp = server.CreateObject("Microsoft.XMLHTTP")
		If IsNull(objHttp) Then
			response.write "서버 연결 오류"
			response.End
		End If
		objHttp.Open "Get", url, False
		objHttp.SetRequestHeader "Content-Type","text/plain"
		objHttp.Send param
		strJsonText = objHttp.responseText
		Set objHttp = Nothing

		getJsonAsp = strJsonText

	End Function

	'// 유저 타임라인 긁어오기 셋팅값
	Sub LoadTweetsUserTimeline()
		' Configure the API call.
		Dim sUsername : sUsername = "your10x10"
		Dim iCount : iCount = 30
		Dim bExcludeReplies : bExcludeReplies = True
		Dim bIncludeRTs : bIncludeRTs = False
		
		Set twSns = objASPTwitter.GetUserTimeline(sUsername, iCount, bExcludeReplies, bIncludeRTs)
	End Sub

	Function IsTweet(ByRef oTweet)
		IsTweet = HasKey(oTweet, "user") 
	End Function

	Function IsRetweet(ByRef oTweet)
		IsRetweet = HasKey(oTweet, "retweeted_status") 
	End Function

	Function IsReply(ByRef oTweet)
		IsReply = Not oTweet.get("in_reply_to_user_id") = Null
	End Function

	Function HasKey(ByRef oTweet, ByVal sKeyName)
		HasKey = Not CStr("" & oTweet.get(sKeyName)) = ""
	End Function

	'// 인코딩
	Function URLsBecomeLinks(sText)
		' Wrap URLs in text with HTML link anchor tags.
		Dim objRegExp
		Set objRegExp = New RegExp
		objRegExp.Pattern = "(http://[^\s<]*)"
		objRegExp.Global = True
		objRegExp.ignorecase = True
		UrlsBecomeLinks = "" & objRegExp.Replace(sText, "")
		Set objRegExp = Nothing
	End Function

	Set objASPTwitter = Nothing
	Set fbSns = Nothing
	Set instaSns = Nothing
	Set twSns = Nothing




	Function getSnsTxtValue()

		Dim SnsTxtVal


		SnsTxtVal ="								<div class='snsTit'> "
		SnsTxtVal = SnsTxtVal & "					<h2><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_tit.png' alt='텐바이텐 소셜 친구들' /></h2> "
		SnsTxtVal = SnsTxtVal & "					<p class='tPad35'><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_subtit.png' alt='텐바이텐 소셜 친구들 소식을 간편하게 한곳에 모아서 즐겨보세요!' /></p> "
		SnsTxtVal = SnsTxtVal & "				</div> "
		SnsTxtVal = SnsTxtVal & "				<p class='rt tMar40' style='padding-right:14px;'><a href='https://www.pinterest.com/your10x10/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_pinterest.png' alt='10x10 핀터레스트도 만나보세요!' /></a></p> "
		SnsTxtVal = SnsTxtVal & "				<div class='snsListWrap'> "
																	Set fbSns = JSON.parse(getJsonAsp("https://graph.facebook.com/181120081908512/posts/?access_token=442566452558488|eOkEkIxvY1hCPmF5Re0jS4LX7XA&limit=8",""))
		SnsTxtVal = SnsTxtVal & "					<!-- facebook --> "
		SnsTxtVal = SnsTxtVal & "					<dl class='facebookfeed'> "
		SnsTxtVal = SnsTxtVal & "						<dt><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_group_fb.png' alt='Facebook' /></dt> "
		SnsTxtVal = SnsTxtVal & "						<dd class='ct' style='padding-bottom:17px;'> "
		SnsTxtVal = SnsTxtVal & "							<iframe src='//www.facebook.com/plugins/like.php?href=https%3A%2F%2Fwww.facebook.com%2Fyour10x10&amp;width=100&amp;layout=button_count&amp;action=like&amp;show_faces=false&amp;share=false&amp;height=21' scrolling='no' frameborder='0' style='border:none; overflow:hidden; width:100px; height:21px;' allowTransparency='true'></iframe> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "						<dd> "
		SnsTxtVal = SnsTxtVal & "							<ul> "
																				If IsNull(fbSns) Or fbSns<>"" Then
																					For i=0 To 10
																						If fbSns.data.Get(i).type="photo" Then
																							j = j+1
		SnsTxtVal = SnsTxtVal & "											<li> "
		SnsTxtVal = SnsTxtVal & "												<div class='box'> "
																									If fbSns.data.Get(i).picture <> "" Then
																										Set fnSnsImg=JSON.parse(getJsonAsp("https://graph.facebook.com/"&fbSns.data.Get(i).id&"/attachments/?access_token=442566452558488|eOkEkIxvY1hCPmF5Re0jS4LX7XA",""))
																										If fnSnsImg.data.Get(0).type="photo" Then
		SnsTxtVal = SnsTxtVal & "															<p class='img'><img src='"&fnSnsImg.data.Get(0).media.image.src&"' alt='' /></p> "
																										ElseIf fnSnsImg.data.Get(0).type="album" Then
		SnsTxtVal = SnsTxtVal & "															<p class='img'><img src='"&fnSnsImg.data.Get(0).subattachments.data.Get(0).media.image.src&"' alt='' /></p> "
																										ElseIf fnSnsImg.data.Get(0).type="cover_photo" Then
		SnsTxtVal = SnsTxtVal & "															<p class='img'><img src='"&fnSnsImg.data.Get(0).media.image.src&"' alt='' /></p> "
																										End If
																										Set fnSnsImg = Nothing
																									End If
		SnsTxtVal = SnsTxtVal & "													<div class='txt'> "
		SnsTxtVal = SnsTxtVal & "														<p>"&URLsBecomeLinks(chrbyte(fbSns.data.Get(i).message, 185, "N"))&" <a href='https://www.facebook.com/your10x10' class='moreView' target='_blank'>... 더보기</a></p> "
		SnsTxtVal = SnsTxtVal & "													</div> "
		SnsTxtVal = SnsTxtVal & "												</div> "
		SnsTxtVal = SnsTxtVal & "											</li> "
																						End If
																						If j>2 Then
																							Exit For 
																						End If
																					Next
																				End If
		SnsTxtVal = SnsTxtVal & "							</ul> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "						<dd class='moreGo'> "
		SnsTxtVal = SnsTxtVal & "							<a href='https://www.facebook.com/your10x10' target='_blank'>페이스북 더보기 &gt;</a> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "					</dl> "

																	'// 트위터 제공키값
																	twitter_api_consumer_key = "1ow2jmpE6U7GytwCEWIGEA"
																	twitter_api_consumer_secret = "Jpldlkh5IXjsiVbxDFZWQny5zJmp0xHoDp6vuBExs"

																	Set objASPTwitter = New ASPTwitter

																	Call objASPTwitter.Configure(TWITTER_API_CONSUMER_KEY, TWITTER_API_CONSUMER_SECRET)
																	Call objASPTwitter.Login
																	Call LoadTweetsUserTimeline
		SnsTxtVal = SnsTxtVal & "					<!-- twitter --> "
		SnsTxtVal = SnsTxtVal & "					<dl class='twitterTimeline'> "
		SnsTxtVal = SnsTxtVal & "						<dt><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_group_tw.png' alt='Twitter' /></dt> "
		SnsTxtVal = SnsTxtVal & "						<dd class='ct' style='padding-bottom:17px;'> "
		SnsTxtVal = SnsTxtVal & "							<a href='https://twitter.com/your10x10' class='twitter-follow-button' data-show-count='false' data-lang='ko'>@your10x10 님 팔로우하기</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+'://platform.twitter.com/widgets.js';fjs.parentNode.insertBefore(js,fjs);}}(document, 'script', 'twitter-wjs');</script>"
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "						<dd> "
		SnsTxtVal = SnsTxtVal & "							<ul> "
																				If IsNull(twSns) Or twSns<>"" Then
																					i = 0
																					For Each oTweet In twSns
																						If oTweet.id_str = "603029422053040128"  Or oTweet.id_str="623033115791921152" Or oTweet.id_str = "623032916323405824" Or oTweet.id_str = "623032180592095232" Or oTweet.id_str = "623031868267466752" Or oTweet.id_str="630984813109886976" Or oTweet.id_str="631632679256084480" Or oTweet.id_str="652057369371742208" Or oTweet.id_str="671950518546128896" Then
																						Else
																							If i > 2 Then
																								Exit For
																							Else
		SnsTxtVal = SnsTxtVal & "												<li> "
		SnsTxtVal = SnsTxtVal & "													<div class='box'> "
																										If Not(IsNull(oTweet.entities.media.Get(0).media_url) Or oTweet.entities.media.Get(0).media_url="") Then
		SnsTxtVal = SnsTxtVal & "															<p class='img'><img src='"&oTweet.entities.media.Get(0).media_url&"' alt='' /></p> "
																										End If
		SnsTxtVal = SnsTxtVal & "														<div class='txt'> "
		SnsTxtVal = SnsTxtVal & "															<p class='wt'><strong class='wtId'>텐바이텐</strong> @your10x10</p> "
		SnsTxtVal = SnsTxtVal & "															<p>"&URLsBecomeLinks(oTweet.text)&" <a href='"&oTweet.entities.media.Get(0).url&"' class='moreView' target='_blank'>... 더보기</a></p> "
		SnsTxtVal = SnsTxtVal & "														</div> "
		SnsTxtVal = SnsTxtVal & "													</div> "
		SnsTxtVal = SnsTxtVal & "												</li> "
																							End If
																						i = i + 1
																						End If
																					Next
																				End If
		SnsTxtVal = SnsTxtVal & "							</ul> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "						<dd class='moreGo'> "
		SnsTxtVal = SnsTxtVal & "							<a href='https://twitter.com/your10x10' target='_blank'>트위터 더보기 &gt;</a> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "					</dl> "

																	Set instaSns = JSON.parse(getJsonAsp("https://api.instagram.com/v1/users/711689678/media/recent/?access_token=711689678.19795ba.a63ebd633d9c4e93a66b25f1e196850f",""))
		SnsTxtVal = SnsTxtVal & "					<!-- instargram --> "
		SnsTxtVal = SnsTxtVal & "					<dl class='instagramTimeline'> "
		SnsTxtVal = SnsTxtVal & "						<dt><img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_group_ig.png' alt='Instargram' /></dt> "
		SnsTxtVal = SnsTxtVal & "						<dd style='padding-top:36px;'> "
		SnsTxtVal = SnsTxtVal & "							<ul> "
																				If IsNull(instaSns) Or instaSns<>"" Then
																					For i=0 To 2
		SnsTxtVal = SnsTxtVal & "										<li> "
		SnsTxtVal = SnsTxtVal & "											<div class='box'> "
		SnsTxtVal = SnsTxtVal & "												<p class='img'><img src='"&instaSns.data.Get(i).images.low_resolution.url&"' alt='' /></p> "
		SnsTxtVal = SnsTxtVal & "												<div class='txt'> "
		SnsTxtVal = SnsTxtVal & "													<p>"&chrbyte(instaSns.data.Get(i).caption.text, 185, "N")&" <a href='"&instaSns.data.Get(i).link&"' class='moreView' target='_blank'>... 더보기</a></p> "
		SnsTxtVal = SnsTxtVal & "												</div> "
		SnsTxtVal = SnsTxtVal & "											</div> "
		SnsTxtVal = SnsTxtVal & "										</li> "
																					Next
																				End If
		SnsTxtVal = SnsTxtVal & "							</ul> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "						<dd class='moreGo'> "
		SnsTxtVal = SnsTxtVal & "							<a href='http://www.instagram.com/your10x10/' target='_blank'>인스타그램 더보기 &gt;</a> "
		SnsTxtVal = SnsTxtVal & "						</dd> "
		SnsTxtVal = SnsTxtVal & "					</dl> "
		SnsTxtVal = SnsTxtVal & "				</div> "
		SnsTxtVal = SnsTxtVal & "				<div class='appDownloadBnr'> "
		SnsTxtVal = SnsTxtVal & "					<img src='http://fiximage.10x10.co.kr/web2013/sns/new_sns_app_bnr.png' alt='텐바이텐 APP을 다운받으시면 더욱 풍성한 소식을 실시간으로 받아보실 수 있습니다.' /> "
		SnsTxtVal = SnsTxtVal & "					<a href='https://itunes.apple.com/kr/app/tenbaiten/id864817011?mt=8' class='btnAppstore' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/sns/sns_btn_appstore.png' alt='App Store' /></a> "
		SnsTxtVal = SnsTxtVal & "					<a href='https://play.google.com/store/apps/details?id=kr.tenbyten.shopping' class='btnGoogle' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/sns/sns_btn_googleplay.png' alt='Google play' /></a> "
		SnsTxtVal = SnsTxtVal & "				</div> "

		getSnsTxtValue = SnsTxtVal

	End Function
	

%>