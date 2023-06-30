var app = angular.module('app', ['ui.router']);

app.config(['$stateProvider', '$urlRouterProvider', function ($stateProvider, $urlRouterProvider) {

    $stateProvider
        .state('Sheet1', {
            url: '/Sheet1',
            templateUrl: 'Sheet1'
        })
        .state('Sheet2', {
            url: '/Sheet2',
            templateUrl: 'Sheet2'
        })
        .state('Sheet3', {
            url: '/Sheet3',
            templateUrl: 'Sheet3'
        })

}]);

app.run(['$http', '$window', function ($http, $window) {
    $http.defaults.headers.common['X-Requested-With'] = 'XMLHttpRequest';
    $http.defaults.headers.common['__RequestVerificationToken'] = $('input[name=__RequestVerificationToken]').val();
}]);

app.service('appService', ['$http', function ($http) {

    // 讀取所有規則
    this.AllRule = function (o) {
        return $http.post('AllRule', o);
    };
    // 取得SQL名稱
    this.GetName = function (o) {
        return $http.post('GetName', o);
    };
    // 讀取SQL資料
    this.GetSQLData = function (o) {
        return $http.post('GetSQLData', o);
    };

}]);

app.factory('dataservice', function () {

    var room = {}

    function set(data) {
        room = data;
    }

    function get() {
        return room;
    }

    return {
        set: set,
        get: get,
    }
});

app.controller('RoomCtrl', ['$scope', '$window', 'appService', '$rootScope', 'dataservice', '$state', function ($scope, $window, appService, $rootScope, dataservice, $state) {

    $scope.data = {
        rule: null,
        allRule: []
    }

    // 讀取SQL資料
    $scope.AllRule = function () {
        appService.AllRule({})
            .then(function (ret) {
                $scope.data.allRule = ret.data;
            })
            .catch(function (ret) {
                //alert('Error');
            });
    }
    $scope.AllRule();

    // 選擇要顯示規則的SQLName
    $scope.GetSQLData = function (sqlName) {
        if (sqlName != undefined) {
            var name_sqlName = sqlName.replace('{', '').replace('}', '').split(",");
            $scope.chooseRule = name_sqlName[0].split(":")[1].replaceAll("\"", ""); // 取得SQL名稱
            sqlName = name_sqlName[1].split(":")[1].replace(/"/g, ""); // <-- 使用正規表示式移除全部的 " , 功能同replaceAll
            // 讀取SQL資料
            appService.GetSQLData({ sqlName: sqlName })
                .then(function (ret) {
                    $scope.SQLData = ret.data;
                    dataservice.set(ret.data)
                    $state.reload(); // 即時更新
                })
                .catch(function (ret) {
                    //alert('Error');
                });
        }
    }
    $scope.GetSQLData();;

}]);

app.controller('Sheet1Ctrl', ['$scope', '$window', 'appService', '$rootScope', '$location', 'dataservice', function ($scope, $window, appService, $rootScope, $location, dataservice) {
    $scope.SQLData = dataservice.get(); // 法規資料
}]);
app.controller('Sheet2Ctrl', ['$scope', '$window', 'appService', '$rootScope', '$location', 'dataservice', function ($scope, $window, appService, $rootScope, $location, dataservice) {
    $scope.SQLData = dataservice.get(); // 法規資料
}]);
app.controller('Sheet3Ctrl', ['$scope', '$window', 'appService', '$rootScope', '$location', 'dataservice', function ($scope, $window, appService, $rootScope, $location, dataservice) {
    $scope.SQLData = dataservice.get(); // 法規資料
}]);