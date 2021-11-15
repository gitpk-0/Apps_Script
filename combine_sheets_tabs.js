function myFunction() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ignoreSheeets = ["MAP", "current for reference", "ALL"];

  const allSheets = ss.getSheets()

  const filteredListOfSheets = allSheets.filter(s => ignoreSheeets.indexOf(s.getSheetName()) == -1)

  let formulaArray = filteredListOfSheets.map(s => `FILTER({'${s.getSheetName()}'!A2:D, "${s.getSheetName()} - Row "&ROW('${s.getSheetName()}'!A2:A)}, '${s.getSheetName()}'!A2:A<>"")`);
  let formulaText = "={" + formulaArray.join(";") + "}";

  //filteredListOfSheets.forEach(s => Logger.log(s.getSheetName()));
  Logger.log(formulaText);

  ss.getSheetByName("ALL").getRange("A2").setFormula(formulaText);
}