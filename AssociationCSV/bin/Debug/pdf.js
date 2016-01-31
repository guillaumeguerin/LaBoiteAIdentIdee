var system = require('system');
var page = new WebPage();
var address = system.args[1];
var exportname = system.args[2]; 

page.paperSize = {
    format        : "A4",
    orientation    : "portrait",
    margin        : { left:"1cm", right:"1cm", top:"1cm", bottom:"1cm" }
};

page.open(address, function (status) {
    page.render(exportname);
    phantom.exit();
});