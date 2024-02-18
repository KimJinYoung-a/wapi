/*

 - version 0.1

�����ϴ� Ű����

 - title, placeholder, maxlength, required, separator
 - input type : tel, number, integer, itemid

 * TODO : confirmAndSave() ��� submit() ����ϴ� ��� ó��
*/

(function($){

    var patternLibrary = {
		tel		: /^([0-9]{0,4}-)?[0-9]{1,4}-[0-9]{4,4}$/,
		number	: /^-?(?:\d+|\d{1,3}(?:,\d{3})+)?(?:\.\d+)?$/,
		integer	: /^-?\d+$/,
		itemid	: /^\d+$/
    };

    var patternMessageLibrary = {
		tel		: "��ȭ��ȣ�� �ùٷ� �Է��ϼ���(��, 000-0000-0000)",
		number	: "���ڸ� �Է°����մϴ�.",
		integer	: "������ �Է°����մϴ�.(�Ҽ��� �ԷºҰ�)",
		itemid	: "�ùٸ� ��ǰ�ڵ尡 �ƴմϴ�."
    };

	// =========================================================================
	// ����Ʈ��
    var defaults = {
        emptyMessage 				: "���� ������ �Է��ϼ���. - #",
		invalidFormatMessage 		: "�߸��� �����Դϴ�.",
		invalidMaxLengthMessage 	: "�ִ� ���̸� �ʰ��Ͽ����ϴ�.(�ִ���� #)",
		confirmMessage				: "�����Ͻðڽ��ϱ�?",
		colorPlaceHolder			: "#AAAAAA",
		colorNormal					: "#000000"
    };
	var opts = null;

    $.fn.enableHTML5 = function(options) {

		opts = $.extend({}, defaults, options);

		// ǰ�� �������� �� �ִ�.
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

			// �Է�â(textarea �� ����)
            $.each($(":input:visible:not(:button, :submit, :radio, :checkbox, select)", form), function(i) {

				// =============================================================
				// placeholder
				// =============================================================

                fillPlaceHolder($(this));

				// �̺�Ʈ ���̱�(�Է�â�� �� ��)
                $(this).bind("focus", function(ev) {
                    ev.preventDefault();
                    if ((this.value == $(this).attr("placeholder")) && $(this).hasClass("classPlaceHolder")) {
                        $(this).attr("value", "");
						$(this).removeClass("classPlaceHolder");
						$(this).css("color", opts.colorNormal);
                    }
                });

				// �̺�Ʈ ���̱�(�Է�â ������ ���� ��)
                $(this).bind("blur", function(ev) {
                    ev.preventDefault();
                    if(this.value == ""){
                        fillPlaceHolder($(this));
                    }
                });

				// =============================================================
				// maxlength
				// TODO : paste �� �̿��� �Է½� ����üũ ����
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

		// ǰ�� �������� �� �ִ�.
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
								// TODO : separator �� �ְ�, ������ ���� ��� üũ
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
					// TODO : required ������ ù��° ������ư�� Ÿ��Ʋ�� ����Ѵ�.
					// TODO : üũ�� ������ ���� ��� ������ư ������ŭ �ߺ�üũ�Ѵ�.
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

			// ǰ�� �������� �� �ִ�.
			$(this).each(function() {

				var form = $(this);

				$.each($(":input:visible:not(:button, :submit, :radio, :checkbox, select)", form), function(i) {
					// placeholder ����
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
