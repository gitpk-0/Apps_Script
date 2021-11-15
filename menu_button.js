function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // menu button name, function button name, function for button to perform
    ui.createMenu('Scripts').addItem('Delete Columns 4 - 9', 'deleteColumns').addToUi();
  }\