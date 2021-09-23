(function () {
    "use strict";
    var placeholderTag = 'n/a';
    var description = 'n/a';
    function GetURLParameter(sParam) {
        var sPageURL = window.location.search.substring(1);
        var sURLVariables = sPageURL.split('&');
        for (var i = 0; i < sURLVariables.length; i++) {
            var sParameterName = sURLVariables[i].split('=');
            if (sParameterName[0] == sParam) {
                return sParameterName[1];
            }
        }
    }
    Office.initialize = function (reason) {
        placeholderTag = decodeURI(GetURLParameter('tag'));
        description = decodeURI(GetURLParameter('description'));
        var holder = document.getElementById("holder");
        holder.innerHTML += '<h3 style="text-align: center;">' + placeholderTag + '</h3>';
        holder.innerHTML += '<hr>';
        holder.innerHTML += '<p style="text-align: left;">Tag name is ' + placeholderTag + '.</p>';
        holder.innerHTML += '<p style="text-align: left;">' + description + '</p>';
    };
})();
//# sourceMappingURL=ViewPlaceholder.js.map