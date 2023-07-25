(function($) {
    "use strict";
    var input = $('.validate-input .input100');
    $('.validate-form').on('submit', function() {
        var check = true;
        for (var i = 0; i < input.length; i++) {
            if (validate(input[i]) == false) {
                showValidate(input[i]);
                check = false;
            }
        }
        return check;
    });
    $('.validate-form .input100').each(function() {
        $(this).focus(function() {
            hideValidate(this);
        });
    });
    function validate(input) {
        if ($(input).attr('type') == 'text' || $(input).attr('username') == 'username') {
            if ($(input).val().trim().match(/[Vv]\d{7}/) == null) {
                return false;
            }
        } else {
            if ($(input).val().trim() == '') {
                return false;
            }
        }
        if ($(input).attr('type') == 'password' || $(input).attr('password') == 'password') {
            if ($(input).val().trim().match(/\d{6}/) == null) {
                return false;
            }
        } else {
            if ($(input).val().trim() == '') {
                return false;
            }
        }
    }
    function showValidate(input) {
        var thisAlert = $(input).parent();
        $(thisAlert).addClass('alert-validate');
    }
    function hideValidate(input) {
        var thisAlert = $(input).parent();
        $(thisAlert).removeClass('alert-validate');
    }
}
)(jQuery);
