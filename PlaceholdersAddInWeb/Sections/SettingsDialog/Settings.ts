(function () {
    "use strict";

    var url;
    var port;
    var token;
    if (localStorage.getItem('url') || localStorage.getItem('url') === '') {
        url = localStorage.getItem('url');
    } else {
        localStorage.setItem('url', 'https://localhost/DocBuilder.Api/api/placeholder/all');
        url = localStorage.getItem('url');
    }

    if (localStorage.getItem('port') || localStorage.getItem('port') === '') {
        port = localStorage.getItem('port');
    } else {
        localStorage.setItem('port', '44390');
        port = localStorage.getItem('port');
    }

    if (localStorage.getItem('token') || localStorage.getItem('token') === '') {
        token = localStorage.getItem('token');
    } else {
        localStorage.setItem('token', '');
        token = localStorage.getItem('token');
    }

    Office.initialize = function (reason) {

        $('input[name="url"]').val(url);
        $('input[name="port"]').val(port);
        $('input[name="token"]').val(token);

        $('#SettingsForm').on('submit', function (e) {
            updateSettings();
        });
    };

    function updateSettings() {
        var formUrl = $('#SettingsForm').find('input[name="url"]').val();
        var formPort = $('#SettingsForm').find('input[name="port"]').val();
        var formToken = $('#SettingsForm').find('input[name="token"]').val();

        localStorage.setItem('url', formUrl);
        localStorage.setItem('port', formPort);
        localStorage.setItem('token', formToken);

        Office.context.ui.messageParent(url);
    }
})();
