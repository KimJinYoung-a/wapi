<%
	'// 선택 변수 선언
	dim editImgUrl, editContent
	dim editWidth, editHeight, blnUploadFile, frmNameCont

	if editWidth="" then editWidth="750px"									'에디터 너비
	if editHeight="" then editHeight="450px"								'에디터 높이
	if blnUploadFile="" or isnull(blnUploadFile) then blnUploadFile=true	'추가 파일업로드 사용여부
	if frmNameCont="" then frmNameCont="editor"								'작성내용 폼이름
	if editImgUrl="" then editImgUrl = webImgUrl & "/TabsWebEditor/"		'이미지 업로드 경로
%>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/util.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/prototype.js"></script> 
<script type="text/javascript" src="/TabsWebEditor/fckeditor/fckeditor.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/fileup.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/multimedia.js"></script>
<script type="text/javascript" src="/TabsWebEditor/fckeditor/tabsfileup4.js"></script>
<script type="text/javascript">
	    // TABSFileupX 오브젝트 생성
	    var shtml;
		var filectl;
		var pluginpage = './FileupX/download.htm';
		if (GBrowser.isMSIE)
		{
			shtml = '<OBJECT ID="filectl" width="747" height="155" border="1" '; 
			shtml += 'CLASSID="CLSID:2342E134-C396-43EC-BCB8-13D513BC5FE0" ';
			shtml += 'CODEBASE="./FileupX/tabsfileup4setup.cab">';
			shtml += '<PARAM NAME="mode" VALUE="upload">';
			shtml += '<PARAM NAME="licensekey" VALUE="' + licensekey + '">';
			shtml += '</OBJECT>';
		}
		else
		{			
			shtml = '<div style="border: solid 1px;width:747px"><embed id="filectl" type="application/x-tabsfileup" width="747" height="150" ';
			shtml += 'mode="upload" licensekey="' + licensekey + '" ';
			shtml += 'pluginspage="' + pluginpage + '"/></div>';
		}		

		if (GBrowser.isMSIE == false && ExistPlugin() == false)
		{
			shtml += '<p><div style="background: #ffffaa;padding: 20px"><ul>';
			shtml += '<li>이 콘텐츠를 표시하는데 사용할 플러그인이 없습니다.</li>';
			shtml += '<li>플러그인을 다운로드 하려면 <a href="' + pluginpage + '">[여기]</a>를 선택하십시오.</li>';
			shtml += '<li>설치 페이지로 이동합니다.</li>';
			shtml += '</ul></div>';
		}
		
		// TABSFileupX 이벤트 처리기
	    function filectl_ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize)
		{
			var overallSize = fileSize + totalFileSize;
			if (totalFileCount >= 10 || overallSize >= 20 * 1024 * 1024)
			{
				if (GBrowser.isChrome)
				{
					filectl.Alert('최대 10개, 20MB까지 업로드할 수 있습니다.');
				}
				else
				{
					alert('최대 10개, 20MB까지 업로드할 수 있습니다.');
				}				
				return false;
			}
			return true;
		}

		function filectl_UploadSuccess(response)
		{
		    alert('성공적으로 업로드 되었습니다.');
		}

		function filectl_UploadSuccessObjectMoved(locationURL)
		{
			location.href = locationURL;	
		}

		function filectl_UploadErrorOccurred(errorType, errorCode, errorDesc, response)
		{
			alert('업로드 오류:\ntype=' + errorType + '\ncode=' + errorCode + '\ndesctiption=' + errorDesc);
			alert('서버 응답:\n' + response);
		}

		function filectl_UploadCanceled()
		{
			alert('업로드가 취소되었습니다.');
		}	

		function addFiles()  
		{
			filectl.AddFile();
		}

		function removeFiles()  
		{
			filectl.RemoveFile();
		}

		function removeAllFiles()  
		{
			filectl.RemoveAllFiles();
		}

		function selectAllFiles()  
		{
			filectl.SelectAllFiles();
		}

		function listFiles()  
		{
			alert('파일뷰에 포함되어 있는 모든 파일 정보를 아이터레이션합니다.');
			var i;
			for (i = 0; i < filectl.FileCount; i++)
			{
				var fileinfo = filectl.GetFileInfo(i);
				alert('Path: ' + fileinfo.FilePath + '\nName: ' + fileinfo.FileName + '\nExt: ' + fileinfo.FileExt + '\nSize: ' + fileinfo.FileSize + '\nURL: ' + fileinfo.FileURL + '\nSelected: ' + fileinfo.Selected);
			}
		}

		function setIconsViewStyle()
		{
			filectl.ViewStyle = 'icons';
		}

		function setListViewStyle()
		{
			filectl.ViewStyle = 'list';
		}

		function setDetailsViewStyle()
		{
			filectl.ViewStyle = 'details';
		}
		
	    var oEditor;
	    function FCKeditor_OnComplete(editorInstance)
        {
            oEditor = editorInstance;
        }
	
		function Submit() {
		    filectl.AddHiddenValue('subject', $('subject').value);
		    filectl.AddHiddenValue('<%=frmNameCont%>', oEditor.GetHTML(true));
		    filectl.AddHiddenValue('newMediaFiles', GetNewImgFilesInfo());  
            filectl.StartUpload();
		}	
		
		function GetNewImgFilesInfo() {
		    var retNewFiles = '';
		    g_newMultimediaHash.each(function(pair) {
		        var fileInfo = pair.value;
		        if (retNewFiles != '') 
		            retNewFiles += '||' + fileInfo.url + '|' + fileInfo.oriname + '|' + fileInfo.filesize + '|' + fileInfo.thumbPath + '|' + fileInfo.isOldFile;
		        else
		            retNewFiles = fileInfo.url + '|' + fileInfo.oriname + '|' + fileInfo.filesize + '|' + fileInfo.thumbPath + '|' + fileInfo.isOldFile;
		    }); 
		    return retNewFiles;
		}

	   // 이미지 업로드 결과 처리		             
		function OnCompleteImageUpload(data) {
		    var ret = data.evalJSON();
		    var multimediaInfo = {
                oriname: ret.displayName,
                filesize: ret.size,
                url: ret.uid,
                thumbPath: ret.uid,
                isOldFile: 'false',
                isPhoto: true
            }; 
		    
		var url = '<%=editImgUrl%>';
            oEditor.InsertHtml('<img src="' + url + multimediaInfo.url +'" />'); 
            $('divMultimediaList').AddMultimedia(multimediaInfo, false);
		}	
</script>
<div id="divPreContent" style="display:none;"><%=editContent%></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
<tr>
	<td>	
		<script type="text/javascript">
		     var oFCKeditor = new FCKeditor('<%=frmNameCont%>', '<%=editWidth%>', '<%=editHeight%>');
		     oFCKeditor.Config["CustomConfigurationsPath"] = '/TabsWebEditor/smart/tabsconfig_lite.js';
		    oFCKeditor.BasePath = '/TabsWebEditor/fckeditor/';		    
		     oFCKeditor.ToolbarSet = 'TABSWebEditor'; 		      		     
                   oFCKeditor.Value = document.getElementById("divPreContent").innerHTML;
		     oFCKeditor.Create();
		</script>
	</td>
</tr>
<tr>
	<td>
		<font color="#8c7301">
		※1. 행나누기(&lt;br&gt;) - 엔터 (Enter Key)<br>
		※2. 문단나누기(&lt;p&gt;) - 시프트 + 엔터 (Shift + Enter Key)
		</font>
	</td>
</tr>
<% if blnUploadFile then %>
<tr>
	<td>
        <input type="button" value="파일추가" onclick="addFiles()" class="button" />
		<input type="button" value="파일삭제" onclick="removeFiles()" class="button" />
		<input type="button" value="모든 파일삭제" onclick="removeAllFiles()" class="button" />
        <!--<input type="button" id="save" value="저장하기" onclick="Submit()" class="button"  />-->
	</td>
</tr>
<tr>
	<td>
        <script type="text/javascript">
			document.writeln(shtml);
			filectl = document.getElementById('filectl');
			if (GBrowser.isMSIE)
			{
				shtml = '<sc' + 'ript type="text/javascript" for="filectl" Event="ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize)">filectl_ChangingUploadFile(filePath, fileSize, totalFileCount, totalFileSize);</sc' + 'ript>';
				shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadSuccess(response)">			filectl_UploadSuccess(response);</sc' + 'ript>';
				shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadSuccessObjectMoved(locationURL)">			filectl_UploadSuccessObjectMoved(locationURL);</sc' + 'ript>';
				shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadErrorOccurred(errorType, errorCode, errorDesc, response)">filectl_UploadErrorOccurred(errorType, errorCode, errorDesc, response);</sc' + 'ript>';
				shtml += '<sc' + 'ript type="text/javascript" for="filectl" Event="UploadCanceled()">filectl_UploadCanceled();</sc' + 'ript>';
				document.writeln(shtml);
			}
			else
			{
				filectl.EventChangingUploadFile = 'filectl_ChangingUploadFile';
				filectl.EventUploadSuccess = 'filectl_UploadSuccess';
				filectl.EventUploadSuccessObjectMoved = 'filectl_UploadSuccessObjectMoved';
				filectl.EventUploadErrorOccurred = 'filectl_UploadErrorOccurred';
				filectl.EventUploadCanceled  = 'filectl_UploadCanceled';
			}
			filectl.UploadURL = 'ASP/insert.asp';
			//filectl.CodePage = 65001;
			filectl.CodePage = 949;
		</script>
	</td>
</tr>
<% end if %>
</table>
<div id="divFileList" style="display:none;">
    <input type="hidden" id="newFiles" name="newFiles" />
    <input type="hidden" id="deletedFiles" name="deletedFiles" />
</div>
<div id="divMultimediaList" style="display:none;">
    <input type="hidden" id="newMediaFiles" name="newMediaFiles" />
    <input type="hidden" id="deletedMediaFiles" name="deletedMediaFiles" />
</div>
