$(function() {
	//GNB Control
	$('.gnb li').mouseover(function() {
		$('.gnb li').removeClass('subOn');
		$(this).addClass('subOn');
	});
	$('.gnb li').mouseout(function() {
		$('.gnb li').removeClass('subOn');
	});
	$('.gnb li:last-child .subNavi').css('right', '-1px');

	//LNB toggle Control
	$('.toggle').click(function() {
		$(this).parent().toggleClass('lnbView');
		if($(this).children('span').html()=="´Ý±â") {
			$(this).children('span').html('¿­±â');
		} else {
			$(this).children('span').html('´Ý±â');
		}
		if($('.container').hasClass('lnbView')) {
			$('.wrap, .contSection').css('min-width', '960px');
		} else {
			$('.wrap, .contSection').css('min-width', '1140px');
		}
	});

	//LNB Control
	$('.lnb dl dt').click(function() {
		$('.lnb dl').removeClass('current');
		$(this).parent().addClass('current');
		$('.lnb dl dd').hide();
		$(this).parent().children('dd').show();
	});

	//List table sorting Control
	$('.sorting').click(function() {
		$(this).children('span').toggleClass('sortWay');
	});

	//help layer Control
	$('.helpBox').mouseover(function(){
		$(this).children('dl').children('dd').show();
	});
	$('.helpBox').mouseleave(function(){
		$(this).children('dl').children('dd').hide();
	});
});