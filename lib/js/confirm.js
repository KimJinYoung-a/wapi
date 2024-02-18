//-----------------------Written by Duke Kim ---------------------------

/*
purpose : 전체적인 폼 엘리먼트 유효성 검사
input : 폼 객체
return : 모든 요소에 문제가 없으면 true, 있으면 false
remark : form 원소의 onsumbit 이벤트에 핸들러를 붙여 submit여부을 결정하는 방식으로 사용
				 예) <form name="register" method="post" action="join_joining_p.asp" onsubmit="return validate(this)">
				 이 함수를 사용하기 위해서는 각 엘리먼트의 id 속성을 목적에 맞게 셋팅해야 한다.

				 ---------------------------------ID Properties GuideLine---------------------------------
				 1. 오류에 대한 메시지박스 출력시 ID 프로퍼티를 사용하므로 사용자에게 전달하려는 적절한 이름을 사용한다.
				 2. 제약사항에 대해서는 [off,off,off,off] 와 같이 대괄호내의 네자리 컴마분리형식을 사용한다. 항목사이 공백넣지 말것.
				 3. 필수입력(공백불허용) 제약조건은 첫자리 셋팅 --> [on,off,off,off]
				 4. 숫자만허용 제약조건은 둘째자리 셋팅 -->			[off,on,off,off]
				 5. 최소길이제한 제약조건은 셋째자리 셋팅 -->		[off,off,5,off]
				 6. 최대길이제한 제약조건은 넷째자리 셋팅 -->		[off,off,off,10]
				 7. 모든 조건을 제약할때는 모든자리 셋팅 -->		[on,on,5,10]

				 example)
					<input type="text" id="[on,off,3,8]아이디" name="id">
					위는 필수입력, 3~8자 길이제한을 둔것이다.
					아이디가 공백일 경우 "아이디란이 비었네요" 식으로 메세지박스가 출력된다.
				 ---------------------------------ID Properties GuideLine---------------------------------
*/

function validate(form)
{

	var rtn_blank, rtn_digit, rtn_length;

	rtn_blank = check_form_blank(form);

	//공백 체크
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//숫자만 허용되는 입력폼 체크
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}

	//최소,최대길이 제한 체크
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}

	return true;
}

function validate3(form)
{
	var rtn_blank, rtn_digit, rtn_length;
	rtn_blank = check_form_blank(form);

	//공백 체크
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//숫자만 허용되는 입력폼 체크
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}
	//최소,최대길이 제한 체크
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}

	//이메일주소의 유효여부 체크(@가 포함되어 있는지 확인)
	else if (check_form_email(form.email.value) == false)
		{
		alert('이메일 주소가 유효하지 않습니다.');
		form.email.focus();
		return false;
		}
	//홈페이지 주소의 유효여부 체크(http://가 포함되어 있는지 확인)
	/*
	else if  (check_form_url(form.g_url.value) == false)
		{
		alert('홈페이지 주소가 유효하지 않습니다.')
		form.g_url.focus();
		return false;
		}

	*/

	// 질문의 선택 여부를 체크
	else if (form.question.value =='')
		{
		alert('질문을 선택 하십시요.');

		return false;

		}


	/*
	//데이터를 넘기기전 비활성 폼 엘리먼트를 활성화 해야 데이터가 전달됨
	var i;
	for (i=0; i < form.elements.length; i++)
		form.elements[i].disabled = false;
	*/
	//폼 제출
	return true;
}

function validate2(form)
{
	var rtn_blank, rtn_digit, rtn_length;
	rtn_blank = check_form_blank(form);

	//공백 체크
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//숫자만 허용되는 입력폼 체크
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}
	//최소,최대길이 제한 체크
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}
	//이메일주소의 유효여부 체크(@가 포함되어 있는지 확인)
	else if (check_form_email(form.buy_email.value) == false)
		{
		alert('이메일 주소가 유효하지 않습니다.');
		form.buy_email.focus();
		return false;
		}
	//홈페이지 주소의 유효여부 체크(http://가 포함되어 있는지 확인)
	/*
	else if  (check_form_url(form.g_url.value) == false)
		{
		alert('홈페이지 주소가 유효하지 않습니다.')
		form.g_url.focus();
		return false;
		}

	*/

	return true;
}



/*
purpose : 아이디 중복여부 체크페이지를 새창에 로딩
input : 폼 객체
remark : 아이디란이 비어있으면 오류메시지, 그렇지 않으면 새창 로딩
*/
function checkid( form )
{
	var id;
	id = 	form.id.value;

	if (id == '')
	{
		alert('아이디란이 비어있네요');
		form.id.focus();
	}
	else
	{
		window.open('/lib/searchid.asp?id=' + id, 'searchid', 'width=400,height=230,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
	}
}
function checkid2()
{
	var id;
	id = 	document.forms[0].id.value;

	if (id == '')
	{
		alert('아이디란이 비어있네요');
		form.id.focus();
	}
	else
	{
		window.open('/lib/searchid.asp?id=' + id, 'searchid', 'width=400,height=230,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
	}
}



/*
purpose : 우편번호 검색 및 적용 페이지를 새창에 로딩
input : 폼 객체
*/
function searchzipcode(type)
{
	window.open('/lib/searchzip.asp?target=' + type, 'searchzip', 'width=460,height=250,scrollbars=yes');
}



/*
purpose : 무료 이메일계정 배포사이트를 새창으로 로딩
*/
function getemail()
	{
	window.open('http://register.daum.net/', 'getemail');
	}



/*
purpose : 폼의 엘리먼트중 value 프로퍼티가 빈칸으로 남아있는 것을 체크
input : 폼 객체
return : 빈칸이 있으면 해당 엘리먼트의 컬렉션내 인덱스번호, 없으면 -1
remark : 엘리먼트 id 프로퍼티의 첫번째 제약조건이 on 이면 빈칸을 허용하지 않는 것을 의미
*/
function check_form_blank(form)
	{
	var i, con, id, pos1, pos2, pos3, pos4;
	for (i=0; i < form.elements.length; i++)
		{
		//텍스트와 패스워드, 텍스트에어리어 및 히든 타입에 대해서만
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea' || form.elements[i].type == 'hidden')
			{
			//제약설정부분을 잘라냄
			id = form.elements[i].id;
			pos1 = id.indexOf(']');							//제약조건과 ID를 구분하는 위치 - 1
			pos2 = id.indexOf('[');							//공백제약조건이 시작되는 위치 - 1
			pos3 = id.indexOf(',', pos2 + 1);		//공백제약조건이 끝나는 위치
			pos4 = id.length;										//ID가 끝나는 위치
			con = id.substring(pos2 + 1, pos3);		//alert(con);
			id = id.substring(pos1 + 1, pos4);		//alert(id);

			//제약설정이 되어있고 빈칸이면
			if (con == 'on' && form.elements[i].value == '')
				{
				alert (id + '란을 입력해주세요.');
				return(i);
				}
			}
		}
	return(-1);
	}



/*
purpose : 폼의 엘리먼트중 숫자만 입력받아야 하는 것을 찾아 value 프로퍼티를 체크
input : 폼 객체
return : 숫자범위를 넘어서는 값이 있으면 해당 엘리먼트의 컬렉션내 인덱스번호, 없으면 -1
remark :
*/
function check_form_digit(form)
	{
	var i, con, id, pos1, pos2, pos3, pos4, j, digit;
	for (i=0; i < form.elements.length; i++)
		{
		//텍스트와 패스워드, 텍스트에어리어 타입에 대해서만
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea')
			{
			//제약설정부분을 잘라냄
			id = form.elements[i].id;
			pos1 = id.indexOf(']');						//제약조건과 ID 를 구분하는 위치 - 1
			pos2 = id.indexOf(',');						//숫자제약조건이 시작되는 위치 - 1
			pos3 = id.indexOf(',', pos2 + 1);	//숫자제약조건이 끝나는 위치
			pos4 = id.length;									//ID가 끝나는 위치
			con = id.substring(pos2 + 1, pos3);		//alert(con);
			id = id.substring(pos1 + 1, pos4);		//alert(id);

			//제약설정이 되어 있으면
			if (con == 'on')
				{
				digit = form.elements[i].value;
				for (j=0; j < digit.length; j++)
					if ((digit.charAt(j) * 0 == 0) == false)
						{
						alert(id + '란은 숫자만 허용됩니다.');
						return(i);
						}
				}
			}
		}
	return(-1);
	}




/*
purpose : 폼의 엘리먼트중 최소길이제한과 최대길이제한을 체크
input : 폼 객체
return : 최대 또는 최소길이를 벗어나는 값이 있으면 해당 엘리먼트의 컬렉션내 인덱스번호, 없으면 -1
remark :
*/
function check_form_length(form)
	{
	var i, id, pos1, pos2, pos3, pos4, max, min, length;
	for (i=0; i < form.elements.length; i++)
		{
		//텍스트와 패스워드, 텍스트에어리어 타입에 대해서만
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea')
			{
			//제약설정부분을 잘라냄
			id = form.elements[i].id;
			pos1 = id.indexOf(']');						//제약조건과 ID 를 구분하는 위치 - 1
			pos2 = id.indexOf(',');						//공백제약조건 건너뜀
			pos2 = id.indexOf(',', pos2 + 1);	//숫자제약조건 건너뜀
			pos2 = id.indexOf(',', pos2);			//최소길이제약조건이 시작되는 위치 - 1
			pos3 = id.indexOf(',', pos2 + 1);	//최소길이제약조건이 끝나는 위치
			min  = id.substring(pos2 + 1, pos3);		//alert(min);
			pos2 = id.indexOf(',', pos2 + 1);	//최대길이제약조건이 시작되는 위치 - 1
			pos3 = id.indexOf(']', pos2);			//최대길이제약조건이 끝나는 위치
			max  = id.substring(pos2 + 1, pos3);		//alert(max);
			pos4 = id.length;									//ID가 끝나는 위치
			id = id.substring(pos1 + 1, pos4);			//alert(id);

			length = (form.elements[i].value).length
			if (id!='' && min != 'off' && max != 'off' )	//최소 또는 최대길이가 설정되어 있는 경우
				{
				if (max == 'off')
					{
					if (min >= length)		//최소길이 제약조건을 위반
						{
						alert(id + '란은 최소한 ' + min + '자 이상이어야 합니다.');
						return(i);
						}
					}
				else if (min == 'off')
					{
					if (length > max)		//최대길이 제약조건을 위반
						{
						alert(id + '란은 최대 ' + max + '자까지만 입력가능합니다.');
						return(i);
						}
					}
				else
					{
					if (min > length || max < length)		//최소, 최대길이 제약조건을 위반
						{
						alert(id + '란은 ' + min + '~' + max + '자 사이로 입력하셔야 합니다.');
						return(i);
						}
					}
				}
			}
		}
	return(-1);
	}




/*
purpose : 주민등록번호의 유효여부 체크
input : 주민번호(하이픈없이 붙여서)
return : 올바르면 true, 올바르지 않으면 false
remark : 바로 밑의 jumin_chk 함수가 반드시 필요하다.
*/

function check_form_ssn(it1, it2){
	var forigndigit = it2.substring(0,1);
	jumin=it1+it2;

	if ((forigndigit=="5")||(forigndigit=="6")){
		return isRegNo_fgnno(jumin);
	}else{
		if(jumin_chk(jumin)){
			return false;
		}else	{
			return true;
		}
	}
}

function isRegNo_fgnno(fgnno) {
        var sum=0;
        var odd=0;
        buf = new Array(13);
        for(i=0; i<13; i++) { buf[i]=parseInt(fgnno.charAt(i)); }
        odd = buf[7]*10 + buf[8];
        if(odd%2 != 0) { return false; }
        if( (buf[11]!=6) && (buf[11]!=7) && (buf[11]!=8) && (buf[11]!=9) ) {
                return false;
        }
        multipliers = [2,3,4,5,6,7,8,9,2,3,4,5];
        for(i=0, sum=0; i<12; i++) { sum += (buf[i] *= multipliers[i]); }
        sum = 11 - (sum%11);
        if(sum >= 10) { sum -= 10; }
        sum += 2;
        if(sum >= 10) { sum -= 10; }
        if(sum != buf[12]) { return false }
        return true;
}

function jumin_chk(it)
{
	IDtot = 0;
	IDAdd="234567892345";

	for(i=0;i<12;i++)
	{
		IDtot=IDtot+parseInt(it.substring(i,i+1))*parseInt(IDAdd.substring(i,i+1));
	}

	IDtot=11-(IDtot%11);
	if(IDtot==10)
	{
		IDtot=0;
	}
	else if(IDtot==11)
	{
		IDtot=1;
	}

	if(parseInt(it.substring(12,13))!=IDtot)
		return true;
	else
		return false;
}




/*
purpose : 이메일 주소의 유효여부 체크
input : 이메일 주소
return : 올바르면 true, 올바르지 않으면 false
remark : 주소에 @가 포함되어 있는지, 또는 두번이상 포함되지는 않았는지 확인
*/

function check_form_email(email)
{

	var pos;


	pos = email.indexOf('@');

	if (pos < 0)				//@가 포함되어 있지 않음
		return(false);
	else
		{
		pos = email.indexOf('@', pos + 1)
		if (pos >= 0)			//@가 두번이상 포함되어 있음
			return(false);
		}


	pos = email.indexOf('.');

	if (pos < 0)				//@가 포함되어 있지 않음
		return false;


	return(true);

}





/*
purpose : URL의 유효여부 체크
input : URL
return : 올바르면 true, 올바르지 않으면 false
remark : 주소가 http://로 시작하는지 확인
*/

function check_form_url(url)
	{
	var protocol;
	protocol = url.substring(0, 7)

	if (protocol != 'http://')				//http://로 시작하지 않음
		return(false);
	else
		return(true);
	}

