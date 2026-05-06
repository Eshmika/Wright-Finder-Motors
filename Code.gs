function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Wright Finder Motors")
    .addItem("Open App", "openApp")
    .addToUi();
}
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Wright Finder Motors App")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function openApp() {
  var html = HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Wright Finder Motors App")
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Wright Finder Motors");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
