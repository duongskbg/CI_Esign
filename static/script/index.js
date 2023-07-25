$(document).ready(function () {
    $(":button").click(function () {
        var api_name = $(this).attr("name");
        $.get("../"+api_name,function (data, status) {
            if (status == "success")
                alert(data);
        })
    });
});