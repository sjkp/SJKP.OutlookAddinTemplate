/// <reference path="../../references.ts" />

module SJKP.OutlookAddin {
    "use strict";

    export class HomeController {
        public static $inject = ["$scope"];
        constructor(private $scope: ng.IScope) {
            Office.initialize = () => {
                console.log('Office context initialized');
            };
            
        }
        
    
    }

    app.controller("homeController", HomeController);
    
}