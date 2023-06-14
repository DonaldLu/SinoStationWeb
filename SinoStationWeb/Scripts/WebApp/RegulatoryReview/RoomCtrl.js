var app = angular.module('app', []);

app.run(['$http', '$window', function ($http, $window) {
    $http.defaults.headers.common['X-Requested-With'] = 'XMLHttpRequest';
    $http.defaults.headers.common['__RequestVerificationToken'] = $('input[name=__RequestVerificationToken]').val();
}]);

app.service('appService', ['$http', function ($http) {

}]);

app.controller('UploadCtrl', ['$scope', '$window', 'appService', '$rootScope', function ($scope, $window, appService, $rootScope) {

    // 開啟關閉上傳按鈕
    $(function () {
        $('#files').change(function () {

            if ($('#files')[0].files.length > 0) {
                $('#uploadBtn').removeAttr('disabled');
            }
            else {
                $('#uploadBtn').attr('disabled', true)
            }
        })
    })

    // 檔案上傳成功或失敗訊息
    $(function () {
        $("#UploadForm").ajaxForm({
            iframe: true,
            dataType: "json",
            success: function (result) {
                $("#UploadForm").resetForm();
                if (result.success) {
                    toastr.success(result.message, 'Success Message')
                }
                else {
                    toastr.error(result.message, 'Error Message')
                }
            },
            error: function (xhr, textStatus, errorThrown) {
                $("#UploadForm").resetForm();
                toastr.error('檔案上傳錯誤.', 'Error Message')
            }
        });
    });

    // 上傳員工名單
    $(document).on("click", "#btnUpload", function () {
        var files = $("#importFile").get(0).files;

        var formData = new FormData();
        formData.append('importFile', files[0]);

        $.ajax({
            url: '/Member/ImportFile',
            data: formData,
            type: 'POST',
            contentType: false,
            processData: false,
            success: function (data) {
                if (data.length > 0) {
                    // 取得所有成員名單
                    //$("#result").html(data);
                    //$("#result").html('<font color="#ff0000">' + data + '</font>');
                    $scope.Test = data;
                    $("#result").html('<div class="row"><div class="col" style="align-items:center" ng-repeat="name in Test"><h6 class="list-group-item" style="color:crimson">{{ name.Name }}</h6></div></div>');
                } else {
                    alert("上傳檔案格式錯誤");
                }
            }
        });

    });

}]);