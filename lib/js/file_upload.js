/*==============================================================*/
/*           탭스업로드 액티브X 컨트롤 함수 start               */
/*==============================================================*/


// 파일 선택 대화상자를 오픈해서 업로드할 파일을 선택한다.
function addFiles()  
{
    frm_upload.TABSFileup.AddFile();
}

// 선택한 파일을 아이콘 뷰에서 제거한다.
function removeFiles()  
{
    frm_upload.TABSFileup.RemoveFile();
}

// 업로드할 파일 목록을 참조한다.
function listFiles()  
{
	var UploadFiles = frm_upload.TABSFileup.UploadFiles;
	var i;
	for (i = 1; i <= UploadFiles.Count; i++) {
		alert("파일 이름: " + UploadFiles(i).SourceFile + "\n\n파일 크기: " + UploadFiles(i).FileSize + " 바이트");
	}
}

// submitForm()의 결과로 발생하는 이벤트
function OnCompletedPostMultipartFormData(ErrType, ErrCode, ErrText, retURL)
{
	if (ErrType == 0) {
		// 오류가 발생하지 않을 경우 성공 페이지로 전환한다.
		location.href = retURL
		//alert("웹설명:" + frm_upload.TABSFileup.Response);
		
    } else {
		// 오류 정보를 간단하게 출력한다.		
		alert("오류형식:" + ErrType + "   오류코드:" + ErrCode + "   오류설명:" + ErrText);
		alert("웹설명:" + frm_upload.TABSFileup.Response);
    }		
		
}

// 아이콘 뷰 모양을 차례로 변경한다.
function changeViewStyle()
{
	var ViewStyle = frm_upload.TABSFileup.ViewStyle;
	if (++ViewStyle > 3) ViewStyle = 0;
	frm_upload.TABSFileup.ViewStyle = ViewStyle;
}


// 컴퍼넌트 인베드 //
function TabsEmbed(md,fid,wd,ht,tg,ft,vs,bg)
{
	switch(md)
	{
		case "view" :
			document.write('<OBJECT ID="' + fid + '" width=' + wd + ' height=' + ht + ' CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78"  CODEBASE="/lib/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="Mode" VALUE="3">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '"> ');
			document.write('</OBJECT>');
			break;

		case "write" :
			document.write('<OBJECT ID="' + fid + '" width="' + wd + '" height="' + ht + '" border=0 CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78" CODEBASE="/lib/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="UploadURL" VALUE="' + tg + '">');
			document.write('<PARAM NAME="CodePage" VALUE="949">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="Mode" VALUE="1">');
			document.write('<PARAM NAME="FileFilter" VALUE="' + ft + '">');
			document.write('</OBJECT>');
			break;

		case "modi" :
			document.write('<OBJECT ID="' + fid + '" width="' + wd + '" height="' + ht + '" border=0 CLASSID="CLSID:BCFA4759-1193-4EC3-92A0-F03F6461DA78" CODEBASE="/util/Tabs/TABSFileup.cab#version=2,1,0,15">');
			document.write('<PARAM NAME="UploadURL" VALUE="' + tg + '">');
			document.write('<PARAM NAME="CodePage" VALUE="949">');
			document.write('<PARAM NAME="ViewStyle" VALUE="' + vs + '">');
			document.write('<PARAM NAME="UploadStyle" VALUE="1">');
			document.write('<PARAM NAME="BkColor" VALUE="' + bg + '">');
			document.write('<PARAM NAME="Key" VALUE="D5C0D007C92B0E978C352FD326461664398A221751847865">');
			document.write('<PARAM NAME="Mode" VALUE="2">');
			document.write('<PARAM NAME="FileFilter" VALUE="' + ft + '">');
			document.write('</OBJECT>');
			break;
	}
}

function OnChangingUploadFile(TotalCount, TotalFileSize)
{
	var UploadFiles = frm_upload.TABSFileup.UploadFiles;
	var i, ft;

    if (TotalCount > 10) {
        alert("너무 많은 파일을 선택하셨습니다.\n\n※ 파일은 10개까지 동시에 업로드 하실 수 있습니다.");
        frm_upload.TABSFileup.StopUpload = true;
    }
	else if ((TotalFileSize/TotalCount) > (2 * 1024 * 1024))
	{
        alert("너무 용량이 큰파일을 선택하셨습니다.\n\n※ 평균 2MB 이하의 용량으로 리사이즈하여 주십시요.");
        frm_upload.TABSFileup.StopUpload = true;
	}
}
