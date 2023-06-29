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
    // 讀取SQL資料
    this.GetSQLData = function (o) {
        return $http.post('GetSQLData', o);
    };

}]);

app.factory('dataservice', function () {

    var room = {}

    function set(data) {
        room = data;
        //room.id = data.id; // ID
        //room.code = data.code; // 代碼
        //room.classification = data.classification; // 區域
        //room.level = data.level; // 樓層
        //room.name = data.name; // 空間名稱(中文)
        //room.engName = data.engName;  // 空間名稱(英文)
        //room.otherName = data.otherName; // 其他名稱
        //room.system = data.system; // 設備/系統
        //room.count = data.count; // 數量
        //room.maxArea = data.maxArea; // 最大面積(m2)
        //room.minArea = data.minArea; // 最小面積(m2)
        //room.demandArea = data.demandArea; // 需求面積
        //room.permit = data.permit; // 容許差異(±%)
        //room.specificationMinWidth = data.specificationMinWidth; // 最小規範寬度
        //room.demandMinWidth = data.demandMinWidth; // 最小需求寬度
        //room.unboundedHeight = data.unboundedHeight; // 規範淨高
        //room.demandUnboundedHeight = data.demandUnboundedHeight; // 需求淨高
        //room.door = data.door; // 門
        //room.doorWidth = data.doorWidth; // 門寬(mm)
        //room.doorHeight = data.doorHeight; // 門高(mm)
    }

    function get() {
        return room;
    }

    return {
        set: set,
        get: get,
    }
});

app.controller('RoomCtrl', ['$scope', '$window', 'appService', '$rootScope', 'dataservice', function ($scope, $window, appService, $rootScope, dataservice) {

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
        appService.GetSQLData({ sqlName: sqlName })
            .then(function (ret) {
                $scope.chooseRule = sqlName;
                $scope.SQLData = ret.data;
                dataservice.set(ret.data)
            })
            .catch(function (ret) {
                //alert('Error');
            });
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