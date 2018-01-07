RegExp.quote = function(str) {
    return (str+'').replace(/[.?*+^$[\]\\(){}|-]/g, "\\$&");
};

document.addEventListener("DOMContentLoaded", function(event) {

    // Pre-kompilar regexi

    var reguliListo = [];
    for (var key in simplaReguli) {
        if (simplaReguli.hasOwnProperty(key)) {
            // Nur remplasar ye vortlimiti
            var re = new RegExp("\\b" + RegExp.quote(key) + "\\b");
            reguliListo.push({de: re, a: simplaReguli[key]});
        }
    }
    console.log(reguliListo.length);
    reguliListo.push.apply(reguliListo, regexalaReguli);
    console.log(reguliListo.length);

    document.getElementById("remplasez").onclick = function() {

        var texto = document.getElementById("fonto").value;
        for (var regulo of reguliListo) {
            texto = texto.replace(regulo.de, regulo.a);
            console.log(regulo.de)
        }

        document.getElementById("rezulto").value = texto;
    }

});