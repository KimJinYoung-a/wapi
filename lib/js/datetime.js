/* ============================================================================
*
* 날짜관련 자바스크립트 공통함수
*
* 날짜를 문자열로 변환하거나 그 반대는
* 문자열이 "2010-01-01 13:21:33" 또는 "2010-01-01" 인 것으로 가정한다.
*
============================================================================ */

/* ============================================================================
유효성 체크

유효한(존재하는) 월(月)인지 체크
function isValidMonth(mm);

유효한(존재하는) 일(日)인지 체크
function isValidDay(yyyy, mm, dd);

유효한(존재하는) 시(時)인지 체크
function isValidHour(hh);

유효한(존재하는) 분(分)인지 체크
function isValidMin(mi);

유효한(존재하는) 초인지 체크
function isValidSec(sec);

============================================================================ */



/* ============================================================================
변환

자바스크립트 Date 객체를 스트링으로 변환
(입력 : Date)
(출력 : "YYYY-MM-DD")
function toDateString(Date date);

자바스크립트 Date 객체를 스트링으로 변환
(입력 : Date)
(출력 : "YYYY-MM-DD HH:MM:SS")
function toTimeString(Date date);

스트링을 자바스크립트 Date 객체로 변환
(입력 : "YYYY-MM-DD HH:MM:SS" 또는 "YYYY-MM-DD")
(출력 : Date)
function toDate(time);

============================================================================ */



/* ============================================================================
날짜 연산

날짜를 더한다.
(입력 : Date)
(출력 : Date)
function addDate(datefrom, day)

두 날짜 사이의 날짜 차이를 구한다.
(입력 : Date)
function getDayInterval(time1,time2)

年을 YYYY형식으로 리턴
(입력 : Date)
function getYear(time)

月을 MM형식으로 리턴
(입력 : Date)
function getMonth(time)

日을 DD형식으로 리턴
(입력 : Date)
function getDay(time)

특정 날짜의 요일
(입력 : Date)
(출력 : '일','월','화','수','목','금','토')
function getDayOfWeek(time)

============================================================================ */



// ============================================================================
function isValidMonth(mm) {
	var m = parseInt(mm,10);
	return (m >= 1 && m <= 12);
}

function isValidDay(yyyy, mm, dd) {
	var m = parseInt(mm,10) - 1;
	var d = parseInt(dd,10);

	var end = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
	if ((yyyy % 4 == 0 && yyyy % 100 != 0) || yyyy % 400 == 0) {
		end[1] = 29;
	}

	return (d >= 1 && d <= end[m]);
}

function isValidHour(hh) {
	var h = parseInt(hh,10);
	return (h >= 1 && h <= 24);
}

function isValidMin(mi) {
	var m = parseInt(mi,10);
	return (m >= 1 && m <= 60);
}

function isValidSec(sec) {
	var s = parseInt(sec,10);
	return (s >= 1 && s <= 60);
}



// ============================================================================
function toDateString(date) {
	var year  = date.getFullYear();
	var month = date.getMonth() + 1; // 1월=0,12월=11이므로 1 더함
	var day   = date.getDate();

	if (("" + month).length == 1) { month = "0" + month; }
	if (("" + day).length   == 1) { day   = "0" + day;   }

	return ("" + year + "-" + month + "-" + day);
}

function toTimeString(date) {
	var hour  = date.getHours();
	var min   = date.getMinutes();
	var sec   = date.getSeconds();

	if (("" + hour).length  == 1) { hour  = "0" + hour;  }
	if (("" + min).length   == 1) { min   = "0" + min;   }
	if (("" + sec).length   == 1) { sec   = "0" + sec;   }

	return ("" + toDateString(date) + " " + hour + ":" + min + ":" + sec)
}

function toDate(time) {
	// (입력 : "YYYY-MM-DD HH:MM:SS" 또는 "YYYY-MM-DD")

	var year, month, day, hour, min, sec;

	if (time.length >= 10) {
		year  = time.substr(0,4);
		month = time.substr(5,2) - 1; // 1월=0,12월=11
		day   = time.substr(8,2);
	}

	if (time.length >= 19) {
		hour  = time.substr(12,2);
		min   = time.substr(15,2);
		sec   = time.substr(18,2);

		return new Date(year,month,day,hour,min, sec);
	}

	return new Date(year,month,day);
}



// ============================================================================
function addDate(datefrom, day) {
	// var dateto = datefrom; <=========== 뻘짓하지말자. 포인터이다.
	var dateto = new Date(datefrom);

	dateto.setDate(datefrom.getDate() + day);

	return dateto;
}

function getDayInterval(date1,date2) {
	var day   = 1000 * 3600 * 24; // 24시간

	return parseInt((date2 - date1) / day, 10);
}

function getYear(time) {
   return time.getFullYear();
}

function getMonth(time) {
   var month = time.getMonth() + 1; // 1월=0,12월=11이므로 1 더함
   if (("" + month).length == 1) { month = "0" + month; }

   return month;
}

function getDay(time) {
   var day = time.getDate();
   if (("" + day).length == 1) { day = "0" + day; }

   return day;
}

function getDayOfWeek(time) {
   var day = time.getDay(); //일요일=0,월요일=1,...,토요일=6
   var week = new Array('일','월','화','수','목','금','토');

   return week[day];
}
