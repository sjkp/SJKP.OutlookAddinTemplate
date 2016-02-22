module SJKP.OutlookAddin {
    'use strict'

    export var app = angular.module('outlookaddin', ['ngRoute', 'AdalAngular', 'officeuifabric.core', 'officeuifabric.components']);
        
    app.config(['$logProvider', function ($logProvider) {
        // set debug logging to on
        if ($logProvider.debugEnabled) {
            $logProvider.debugEnabled(true);
        }
    }]);
    app.constant('sharePointUrl', 'https://henrik5.sharepoint.com'); //'https://delegatedemo01.sharepoint.com'
    app.constant('crmUrl', 'https://henrik5.crm4.dynamics.com');
    app.constant('appClientId', '05b5d59f-0cc2-4d7c-bd98-04321d87e3f2'); //'68678bcc-397c-4261-b9e5-e0d3d981f7f2'

    if (window.location.host == 'localhost') {
        app.constant('backendUrl', 'https://localhost:44302/api');
    }
    else {
        app.constant('backendUrl', '/api');
    }
    app.config(Config);

}