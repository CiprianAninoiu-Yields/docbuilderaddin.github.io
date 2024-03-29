﻿(function () {
    "use strict";

    var url;
    var token;
    if (localStorage.getItem('url') || localStorage.getItem('url') === '') {
        url = localStorage.getItem('url');
    } else {
        localStorage.setItem('url', `https://localhost/CE.ApiGateway/api/DocBuilder/placeholders`);
        url = localStorage.getItem('url');
    }

    if (localStorage.getItem('token') || localStorage.getItem('token') === '') {
        token = localStorage.getItem('token');
    } else {
        localStorage.setItem('token', '1234');
        token = localStorage.getItem('token');
    }

    Office.initialize = function (reason) {

        $('input[name="url"]').val(url);
        $('input[name="token"]').val(token);

        $('#SettingsForm').on('submit', function (e) {
            updateSettings();
        });
    };

    function updateSettings() {
        var formUrl = $('#SettingsForm').find('input[name="url"]').val();
        var formToken = $('#SettingsForm').find('input[name="token"]').val();

        localStorage.setItem('url', formUrl);
        localStorage.setItem('token', formToken);

        Office.context.ui.messageParent(url);
    }
})();
