/*

 - version 0.1

대응하는 키워드

 - title, placeholder, maxlength, required, separator
 - input type : tel, number, integer, itemid

 * TODO : confirmAndSave() 대신 submit() 사용하는 경우 처리
*/

(function($){

    var patternLibrary = {
		tel		: /^([0-9]{0,4}-)?[0-9]{1,4}-[0-9]{4,4}$/,
		number	: /^-?(?:\d+|\d{1,3}(?:,\d{3})+)?(?:\.\d+)?$/,
		integer	: /^-?\d+$/,
		itemid	: /^\d+$/
    };

    var patternMessageLibrary = {
		tel		: "전화번호를 올바로 입력하세요(예, 000-0000-0000)",
		number	: "숫자만 입력가능합니다.",
		integer	: "정수만 입력가능합니다.(소수점 입력불가)",
		itemid	: "올바른 상품코드가 아닙니다."
    };

	// =========================================================================
	// 디폴트값
    var defaults = {
        emptyMessage 				: "다음 내용을 입력하세요. - #",
		invalidFormatMessage 		: "잘못된 형식입니다.",
		invalidMaxLengthMessage 	: "최대 길이를 초과하였습니다.(최대길이 #)",
		confirmMessage				: "저장하시겠습니까?",
		colorPlaceHolder			: "#AAAAAA",
		colorNormal					: "#000000"
    };
	var opts = null;

    $.fn.enableHTML5 = function(options) {

		opts = $.extend({}, defaults, options);

		// 품은 여러개일 수 있다.
		$(this).each(function() {

            var form = $(this);

            function fillPlaceHolder(input) {
				if (input.attr("placeholder") && input.attr("placeholder") != "" && input.attr("type") != "password") {
					if (!input.attr("value") || input.attr("value") == "") {
						input.val(input.attr("placeholder"));
						input.css("color", opts.colorPlaceHolder);
						input.addClass("classPlaceHolder");
					}
				}
            }

			// 입력창(textarea 도 포함)
            $.each($(":input:visible:not(:button, :submit, :radio, :checkbox, select)", form), function(i) {

				// =============================================================
				// placeholder
				// =============================================================

                fillPlaceHolder($(this));

				// 이벤트 붙이기(입력창에 들어갈 때)
                $(this).bind("focus", function(ev) {
                    ev.preventDefault();
                    if ((this.value == $(this).attr("placeholder")) && $(this).hasClass("classPlaceHolder")) {
                        $(this).attr("value", "");
						$(this).removeClass("classPlaceHolder");
						$(this).css("color", opts.colorNormal);
                    }
                });

				// 이벤트 붙이기(입력창 밖으로 나갈 때)
                $(this).bind("blur", function(ev) {
                    ev.preventDefault();
                    if(this.value == ""){
                        fillPlaceHolder($(this));
                    }
                });

				// =============================================================
				// maxlength
				// TODO : paste 를 이용한 입력시 길이체크 않함
				// =============================================================

                $("textarea").filter(this).each(function() {
                    if ($(this).attr("maxlength") > 0) {
                        $(this).keypress(function(ev) {
                            var cc = ev.charCode || ev.keyCode;
                            if(cc == 37 || cc == 39) {
                                return true;
                            }
                            if (cc == 8 || cc == 46) {
                                return true;
                            }
                            if (this.value.length >= $(this).attr("maxlength")){
                                return false;
                            } else {
                                return true;
                            }
                        });

                    }
                });

			});
		});

		return $(this);
	}
    $.fn.validate = function() {

		var result = true;

		// 품은 여러개일 수 있다.
		$(this).each(function() {

			if (result != true) {
				return false;
			}

            var form = $(this);

			$.each($(":input:visible:not(:button, :submit, :radio, :checkbox, select)", form), function(i) {

				if (result != true) {
					return false;
				}

				// =============================================================
				// maxlength
				// =============================================================

				if ($(this).attr("maxlength")) {
					if (($(this).hasClass("classPlaceHolder") != true) && (this.value.length > $(this).attr("maxlength"))) {
						alert(opts.invalidMaxLengthMessage.replace("#", $(this).attr("maxlength")));
						this.focus();
						result = false;
						return false;		// break each()
					}
				}

				// =============================================================
				// required - input, textarea
				// =============================================================

				if ($(this).attr("required")) {
					if (($(this).hasClass("classPlaceHolder") == true) || (this.value == "")) {
						alert(opts.emptyMessage.replace("#", $(this).attr("title")));
						this.focus();
						result = false;
						return false;		// break each()
					}
				}

				// =============================================================
				// custom pattern
				// =============================================================

				if ($(this).attr("type")) {
					if (($(this).hasClass("classPlaceHolder") != true) && (this.value.length > 0)) {
						var thisType = $(this).attr("type");
						var thisPattern = null;
						$.each(patternLibrary, function (key, value) {
							if (thisType == key.toString()) {
								thisPattern = value;
								return false;		// break each()
							}
						});

						if (thisPattern != null) {

							var thisErrorMessage = opts.invalidFormatMessage;
							$.each(patternMessageLibrary, function (key, value) {
								if (thisType == key.toString()) {
									thisErrorMessage = value.toString();
									return false;		// break each()
								}
							});

							if ($(this).attr("separator")) {
								// TODO : separator 만 있고, 내용이 없는 경우 체크
								var arr = this.value.split($(this).attr("separator"));

								for (var i = 0; i < arr.length; i++) {
									if (arr[i] != "") {
										var match = arr[i].match(thisPattern);
										if (!match) {
											alert(thisErrorMessage.replace("#", $(this).attr("title")));
											this.focus();
											result = false;
											return false;		// break each()
										}
									}
								}
								//
							} else {
								var match = this.value.match(thisPattern);
								if (!match) {
									alert(thisErrorMessage.replace("#", $(this).attr("title")));
									this.focus();
									result = false;
									return false;		// break each()
								}
							}
						}
					}
				}

			});

			$.each($("select", form), function(i) {

				if (result != true) {
					return false;
				}

				// =============================================================
				// required - select
				// =============================================================

				if ($(this).attr("required")) {
					if (this.value == "") {
						alert(opts.emptyMessage.replace("#", $(this).attr("title")));
						this.focus();
						result = false;
						return false;		// break each()
					}
				}

			});

			$.each($("input[type=radio]", form), function(i) {

				if (result != true) {
					return false;
				}

				// =============================================================
				// required - radio
				// =============================================================

				if ($(this).attr("required")) {
					// TODO : required 지정된 첫번째 라디오버튼의 타이틀을 사용한다.
					// TODO : 체크에 에러가 없는 경우 라디오버튼 갯수만큼 중복체크한다.
					var thisTitle = $(this).attr("title");
					var thisObject = this;
					if ($("input[type=radio][name=" + $(this).attr("name") + "]:checked", form).val()) {
						//
					} else {
						alert(opts.emptyMessage.replace("#", thisTitle));
						thisObject.focus();
						result = false;
						return false;
					}
				}

			});

		});

		return result;
	}
	$.fn.confirmAndSave = function() {
		if (confirm(opts.confirmMessage) == true) {

			// 품은 여러개일 수 있다.
			$(this).each(function() {

				var form = $(this);

				$.each($(":input:visible:not(:button, :submit, :radio, :checkbox, select)", form), function(i) {
					// placeholder 삭제
                    if ((this.value == $(this).attr("placeholder")) && $(this).hasClass("classPlaceHolder")) {
                        $(this).attr("value", "");
						$(this).removeClass("classPlaceHolder");
						$(this).css("color", opts.colorNormal);
                    }
				});

				this.submit();
			});

		}

		return $(this);
	}
})(jQuery);
