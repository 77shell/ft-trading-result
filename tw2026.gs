var TW2026 = SpreadsheetApp.getActive().getSheetByName('TW.2026');

function updateTW2026_DayChange() {
  var daychange = TW2026.getRange('H2:H14');
  // Calculate total number of cells
  var totalCells = daychange.getNumRows() * daychange.getNumColumns();
  var shares = TW2026.getRange('C2:C14');

  // Tuesday to Satursday cases : "(D2-index(googlefinance(A2, \"price\",(today()-1)),2,2))*C2";
  // Sunday case : "(D2-index(googlefinance(A2, \"price\",(today()-2)),2,2))*C2";
  // Monday case : "(D2-index(googlefinance(A2, \"price\",(today()-3)),2,2))*C2";
  var formulahead = "-index(googlefinance(A";
  var forumlacenter = ", \"price\",(today()-";
  var forumlatail = ")),2,2))*C";

  var day = new Date().getDay();
  var j = 2;
  for (var i = 1; i <= totalCells; ++i, ++j) {
    var c = daychange.getCell(i, 1);
    var s = shares.getCell(i, 1).getValue();
    Logger.log(s); 
    if (s < 1 || s == "") { // 1 : one share
      c.setValue("");
      continue;
    }
    switch (day) {
      case 0: // Sunday (Last price is on Friday, so compare Friday and Thursday{-3})
        c.setFormula('(D' + j + formulahead + j + forumlacenter + '3' + forumlatail + j);
        break;

      case 1: // Monday (Compare Monday and Friday{-3})
        c.setFormula('(D' + j + formulahead + j + forumlacenter + '3' + forumlatail + j);
        break;

      default: // Tuesday ~ Satursday
        c.setFormula('(D' + j + formulahead + j + forumlacenter + '1' + forumlatail + j);
        break;
    }
  }
}

/*
  0	日 - 3
  1	一 - 3
  2	二 - 1
  3	三 - 1
  4	四 - 1
  5	五 - 1
  6	六 - 1

  if (Top9sheet) {
    var tickets = Top9sheet.getRange('C2:C');
    var prices = Top9sheet.getRange('N4:N');
    var j = 0;
    for (var i = 1; i <= N; i++, j++) {
      var p = prices.getCell(j * 2 + 1, 1);
      var t = tickets.getCell(i, 1).getValue();
      // GOOGLEFINANCE(ticket, "price", date, days)
      var formula = "=GOOGLEFINANCE(\"" + t + "\")";
      p.setFormula(formula);
      p.setFontWeight("bold");
      p.setFontColor("blue");
      p.setNumberFormat("0.00");
    }
  }
*/