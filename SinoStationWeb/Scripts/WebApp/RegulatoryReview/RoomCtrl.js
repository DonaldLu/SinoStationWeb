var app = angular.module('app', []);

app.run(['$http', '$window', function ($http, $window) {
    $http.defaults.headers.common['X-Requested-With'] = 'XMLHttpRequest';
    $http.defaults.headers.common['__RequestVerificationToken'] = $('input[name=__RequestVerificationToken]').val();
}]);

app.service('appService', ['$http', function ($http) {

}]);

app.controller('RoomCtrl', ['$scope', '$window', 'appService', '$rootScope', function ($scope, $window, appService, $rootScope) {

    // 上傳Excel檔
    $(document).on("click", "#btnUpload", function () {
        var files = $("#importFile").get(0).files;

        var formData = new FormData();
        formData.append('file', files[0]);

        $.ajax({
            url: '/RegulatoryReview/Upload',
            data: formData,
            type: 'POST',
            contentType: false,
            processData: false,
            success: function (data) {
                if (data.length > 0) {
                    /*$("#result").html('<div class="col-sm-3" style="align-items:center" ng-repeat="item in data"><p> {{item}} </p></div>');*/
                } else {
                    alert("上傳檔案格式錯誤");
                }
            }
        });
    });

}]);