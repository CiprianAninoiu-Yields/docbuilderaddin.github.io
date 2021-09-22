var dialog;
var url;
var port;
var token;
if (localStorage.getItem('url') || localStorage.getItem('url') === '') {
    url = localStorage.getItem('url');
}
else {
    localStorage.setItem('url', 'https://localhost/DocBuilder.Api/api/placeholder/all');
    url = localStorage.getItem('url');
}
if (localStorage.getItem('port') || localStorage.getItem('port') === '') {
    port = localStorage.getItem('port');
}
else {
    localStorage.setItem('port', '44390');
    port = localStorage.getItem('port');
}
if (localStorage.getItem('token') || localStorage.getItem('token') === '') {
    token = localStorage.getItem('token');
}
else {
    localStorage.setItem('token', '');
    token = localStorage.getItem('token');
}
(function () {
    "use strict";
    var messageBanner;
    Office.initialize = function (reason) {
        window.Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            loadSampleData();
            setUpPlaceholders();
        });
    };
    function viewPlaceholderInfo(placeholderTag, description) {
        Office.context.ui.displayDialogAsync('https://cipriananinoiu-yields.github.io/docbuilderaddin.github.io/PlaceholdersAddInWeb/Sections/ViewPlaceholderDialog/ViewPlaceholder.html?tag=' + placeholderTag + '&description=' + description, { height: 20, width: 20 }, function (asyncResult) {
            dialog = asyncResult.value;
        });
    }
    function setUpPlaceholders() {
        var placeholders = new Array();
        getPlaceholders().then(function (response) {
            placeholders = response;
            var holder = document.getElementById("holder");
            if (placeholders.length === 0) {
                holder.innerHTML += '<p style="padding: 3px; margin: 0px;>' + 'No Placeholders Found' + '</p>';
                return;
            }
            var placeholderTypes = placeholders.map(function (x) { return x.docPlaceholderType; });
            var uniquePlTypes = new Array();
            placeholderTypes.forEach(function (x) {
                if (uniquePlTypes.filter(function (y) { return y.id === x.id; }).length === 0)
                    uniquePlTypes.push(x);
            });
            for (var j = 0; j < uniquePlTypes.length; j++) {
                holder.innerHTML += '<p style="background-color:#377dde; padding: 3px; margin: 0px; border-style: groove; border-width: 2px; font-weight: bold">' + uniquePlTypes[j].name + '</p>';
                var typeplaceholders = placeholders.filter(function (x) { return x.docPlaceholderType.id === uniquePlTypes[j].id; });
                typeplaceholders.forEach(function (x) {
                    var placeholder = $('<div/>').attr({
                        style: 'display: flex; justify-content: space-between; padding: 3px; margin: 0px; border-style: groove; border-color:#666; border-width: 1px 2px',
                    });
                    var button = $('<input/>').attr({
                        id: 'placeholder' + x.id,
                        style: 'background: none; border: none; margin: 0; padding: 0; cursor: pointer; margin-left: 5px',
                        type: 'button',
                        name: x.name,
                        value: x.name
                    });
                    var infobox = $('<input/>').attr({
                        id: 'placeholder-info' + x.id,
                        style: 'background: none; margin-right: 10px',
                        type: 'button',
                        name: '+',
                        value: '+'
                    });
                    placeholder.append(button);
                    placeholder.append(infobox);
                    $("#holder").append(placeholder);
                });
            }
            placeholders.forEach(function (x) {
                $("#placeholder" + x.id).click(function () {
                    insertPlaceholder(x.tag);
                });
                $("#placeholder-info" + x.id).click(function () {
                    viewPlaceholderInfo(x.tag, x.description);
                });
            });
        });
    }
    function getPlaceholders() {
        return fetch(url)
            .then(function (res) {
            if (!res.ok) {
                throw new Error("N");
            }
            else {
                return res.json();
            }
        });
    }
    function insertPlaceholder(placeholder) {
        Word.run(function (context) {
            var range = context.document.getSelection();
            context.load(range, 'text');
            range.clear();
            range.delete();
            return context.sync()
                .then(function () {
                range.insertText(placeholder, Word.InsertLocation.end);
            })
                .then(context.sync);
        })
            .catch(errorHandler);
    }
    function loadSampleData() {
        Word.run(function (context) {
            var body = context.document.body;
            body.clear();
            body.insertText("This is a sample text inserted in the document", Word.InsertLocation.end);
            return context.sync();
        })
            .catch(errorHandler);
    }
    function errorHandler(error) {
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
function openSettings(event) {
    Office.context.ui.displayDialogAsync('https://cipriananinoiu-yields.github.io/docbuilderaddin.github.io/PlaceholdersAddInWeb/Sections/SettingsDialog/Settings.html', { height: 30, width: 20 }, function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processSettings);
    });
    event.completed();
}
function processSettings(arg) {
    dialog.close();
}
var Placeholder = /** @class */ (function () {
    function Placeholder() {
    }
    return Placeholder;
}());
var PlaceholderType = /** @class */ (function () {
    function PlaceholderType() {
    }
    return PlaceholderType;
}());
//# sourceMappingURL=Home.js.map