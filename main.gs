/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

var Top9sheet = SpreadsheetApp.getActive().getSheetByName('QQQ Top9');
var N = 8;
/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems_qqq = [
	{ name: 'Download latest QQQ portfilio', functionName: 'downloadLatestQQQ_' },
	{ name: 'Flush QQQ Top9', functionName: 'flush_QQQ_top9_sheet' }
    ];
    spreadsheet.addMenu('QQQ', menuItems_qqq);

    var menuItems_irr = [
	{ name: 'Calculate IRR', functionName: 'calculate_irr' },
	{ name: 'Average IRR', functionName: 'average_irr'}
    ];
    spreadsheet.addMenu('IRR', menuItems_irr);

    var menuItems_2026 = [
	{ name: 'Calculate Gain/Loss', functionName: 'calculate_gainloss' }
    ];
    spreadsheet.addMenu('2026', menuItems_2026);

    // flush_QQQ_top9_sheet();
    // listSheets_();
    updateTW2026_DayChange();
}

function listSheets_() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheetname = 'Summary';
    var sumsheet = spreadsheet.getSheetByName(sheetname);
    if (!sumsheet) {
	sumsheet = spreadsheet.insertSheet(sheetname, spreadsheet.getNumSheets());
    }
    spreadsheet.setActiveSheet(sumsheet);
    spreadsheet.moveActiveSheet(1);

    var range = sumsheet.getRange('B1:B');
    var nbr = spreadsheet.getNumSheets();
    Logger.log('Sheet number' + nbr);
    for (var i = 1; i < nbr; i++) {
	Logger.log(i + spreadsheet.getSheets()[i].getName());
	range.getCell(i, 1).setValue(spreadsheet.getSheets()[i].getName());
    }
}

function downloadLatestQQQ_() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheetname = 'QQQ -> ' + Date();
    var portfolioSheet = spreadsheet.getSheetByName(sheetname);
    if (portfolioSheet) {
	portfolioSheet.clear();
	portfolioSheet.activate();
    }
    else {
	portfolioSheet = spreadsheet.insertSheet(sheetname, spreadsheet.getNumSheets());
    }

    var cell = portfolioSheet.getRange('A1');
    /* portfolioSheet.setCurrentCell(cell); */
    cell.setFormulaR1C1('=IMPORTDATA("https://www.invesco.com/us/financial-products/etfs/holdings/main/holdings/0?audienceType=Investor&action=download&ticker=QQQ")');
    // listSheets_();
}

var Top9_Weigh_Sum;
function flush_QQQ_top9_sheet() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheetname = 'QQQ Top9';
    Top9sheet = spreadsheet.getSheetByName(sheetname);
    if (Top9sheet) {
	Top9sheet.clear();
	Top9sheet.activate();
    }
    else {
	Top9sheet = spreadsheet.insertSheet(sheetname, spreadsheet.getNumSheets());
    }
    spreadsheet.moveActiveSheet(2);
    Top9sheet.getRange('A1').setFormulaR1C1('=IMPORTDATA("https://www.invesco.com/us/financial-products/etfs/holdings/main/holdings/0?audienceType=Investor&action=download&ticker=QQQ")');
    var ticketsrange = Top9sheet.getRange('C2:C9');
    var weightrange = Top9sheet.getRange('F2:F9');
    ticketsrange.setBackground('#ddddaa');
    ticketsrange.setFontWeight("bold");
    weightrange.setBackground('#ddccaa');
    weightrange.setFontWeight("bold");

    // Sum up top9 weigh
    var sumtitle = Top9sheet.getRange('K1');
    sumtitle.setValue("Top9 Weigh:");
    sumtitle.setFontWeight("bold");

    var top9sum = Top9sheet.getRange('L1');
    top9sum.setFormula('=SUM(F2:F9)');
    top9sum.setFontWeight("bold");
    Top9_Weigh_Sum = top9sum.getValue();

    var startdate = "Date(2022,2,25)";
    get_top9_prices(startdate);
    get_top9_today_prices();
    create_lables();
    calculate_top9_share();
    SpreadsheetApp.flush();
}

function get_top9_prices(startdate) {
    // var top9sheet = SpreadsheetApp.getActive().getSheetByName('QQQ Top9');
    if (Top9sheet) {
	var tickets = Top9sheet.getRange('C2:C10');
	var prices = Top9sheet.getRange('L2:L');
	var titles = Top9sheet.getRange('K2:K20');
	var weighs = Top9sheet.getRange('F2:F10');
	var top9weighs = Top9sheet.getRange('O2:O');
	for (var i = 1; i <= N; ++i) {
	    var p = prices.getCell(i * 2, 1);
	    var t = tickets.getCell(i, 1).getValue();
	    var w = top9weighs.getCell(i * 2 + 1, 1);
	    w.setValue(100 / Top9_Weigh_Sum * weighs.getCell(i, 1).getValue());
	    w.setNumberFormat("0.00");
	    
	    var title = titles.getCell(i * 2 + 1, 1);
	    title.setValue(t);
	    title.setFontWeight('bold');
	    // GOOGLEFINANCE(ticket, "price", date, days)
	    var formula = "=GOOGLEFINANCE(\"" + t + "\",\"price\"," + startdate + ",1)";
	    p.setFormula(formula);
	}

	var priceoutput = Top9sheet.getRange('M2:M');
	for (var i = 0; i < 10; i++) {
	    var c = priceoutput.getCell(i * 2 + 1, 1);
	    c.setFontWeight("bold");
	    c.setFontColor("green");
	}
    }
}

function get_top9_today_prices() {
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
}

function create_lables() {
    var row = Top9sheet.getRange('N3:T3');
    row.setFontWeight('bold');
    var labels = [
	['Price', 'Weigh', 'Cost', 'Share', 'Bought share', 'Diff', 'Diff in cash']
    ];
    row.setValues(labels);
}

function calculate_top9_share() {
    var costs = Top9sheet.getRange('P4:P25');
    var shares = Top9sheet.getRange('Q4:Q25');
    var prices = Top9sheet.getRange('N4:N25');
    var weighs = Top9sheet.getRange('O4:O25');
    var tickets = Top9sheet.getRange('K4:K25');
    var bougshares = Top9sheet.getRange('R4:R25');
    var difference = Top9sheet.getRange('S4:S25');
    var diffincash = Top9sheet.getRange('T4:T25');

    var sheet = SpreadsheetApp.getActive(); 
    var capital = sheet.getRange('myqqq-2021-10-11!M2').getValue();
    var cash = sheet.getRange('myqqq-2021-10-11!M7').getValue();
    var exclude = sheet.getRange('myqqq-2021-10-11!M9').getValue();
    capital += cash;
    capital -= exclude;
    var c = Top9sheet.getRange('P2');
    c.setValue(capital);
    c.setNumberFormat("0,000");
    c.setFontWeight('bold');

    for (var r = 1; r <= N * 2; r += 2) {
	var c = costs.getCell(r, 1);
	var p = prices.getCell(r, 1);
	var w = weighs.getCell(r, 1);
	var s = shares.getCell(r, 1);
	var t = tickets.getCell(r, 1);
	var b = bougshares.getCell(r, 1);
	var d = difference.getCell(r, 1);
	var d2 = diffincash.getCell(r,1);
	
	var cost = capital * w.getValue() / 100;
	//c.setValue(cost);
	var formula = '=P2*' + w.getA1Notation() + '/100';
	c.setFormula(formula);
	c.setNumberFormat("$0,000");

	// Logger.log(p.getValue());
	// var share = c.getValue() / p.getValue();
	// s.setValue(share);
	var formu = "=" + c.getA1Notation() + "/" + p.getA1Notation();
	s.setFormula(formu);
	s.setNumberFormat("0.0");

	// Logger.log(t.getValue());
	b.setValue(get_bought_shares(t.getValue()));

	// var diff = s.getValue() - b.getValue();
	formu = "=" + s.getA1Notation() + "-" + b.getA1Notation();
	d.setFormula(formu);
	d.setNumberFormat("0.00");

	// Difference in cash
	formu = '=' + d.getA1Notation() + '*' + p.getA1Notation();
	d2.setFormula(formu);
	d2.setNumberFormat("$000");
    }
}

function get_bought_shares(target) {
    var sumsheet = SpreadsheetApp.getActive().getSheetByName('myqqq-2021-10-11');
    if (sumsheet) {
	var boughtickets = sumsheet.getRange('A2:A12');
	var f = sumsheet.createTextFinder(target);
	var result = f.findNext();

	var r = 'D' + result.getRowIndex();
	var boughtshares = sumsheet.getRange(r).getValue();
	// Logger.log(boughtshares);
	return boughtshares;
    }
}

function average_5y_irr(sheet) {
    var nbr = 0;
    var irr_sum = 0;
    var _5y_irrs = sheet.getRange('AD2:AD22');
    var rows = _5y_irrs.getNumRows();
    // Look the latest 5
    var r = rows - 5 + 1;
    for (; r <= rows; ++r) {
	var cel = _5y_irrs.getCell(r, 1);
	var avg = cel.getValue();
	if (typeof avg == 'number') {
	    Logger.log("r:%s %s", r, avg);
	    ++nbr;
	    irr_sum += avg;
	}
	else {
	    Logger.log("NG r:%s %s", r, avg);
	}
    }
    if (nbr > 0) {
	var _5y_irr_avg = sheet.getRange('AL25');
	_5y_irr_avg.setValue(irr_sum / nbr);
	sheet.getRange('AL26').setValue(irr_sum);
	sheet.getRange('AL27').setValue(nbr);
    }
    else {
	sheet.getRange('AL26').setValue(0);
	sheet.getRange('AL27').setValue(0);
    }
}

function average_10y_irr(sheet) {
    var nbr = 0;
    var irr_sum = 0;
    var _10y_irrs = sheet.getRange('AF2:AF17');
    var rows = _10y_irrs.getNumRows();
    var r = rows - 5 + 1;
    for (; r <= rows; ++r) {
	var cel = _10y_irrs.getCell(r, 1);
	var avg = cel.getValue();
	if (typeof avg == 'number') {
	    ++nbr;
	    irr_sum += avg;
	}
    }
    if (nbr > 0) {
	var _10y_irr_avg = sheet.getRange('AN25');
	_10y_irr_avg.setValue(irr_sum / nbr);
	sheet.getRange('AN26').setValue(irr_sum);
	sheet.getRange('AN27').setValue(nbr);
    }
}

function average_15y_irr(sheet) {
    var nbr = 0;
    var irr_sum = 0;
    var _15y_irrs = sheet.getRange('AH2:AH12');
    var rows = _15y_irrs.getNumRows();
    for (var r = 1; r <= rows; ++r) {
	var cel = _15y_irrs.getCell(r, 1);
	var avg = cel.getValue();
	if (typeof avg == 'number') {
	    ++nbr;
	    irr_sum += avg;
	}
    }
    if (nbr > 0) {
	var _15y_irr_avg = sheet.getRange('AP25');
	_15y_irr_avg.setValue(irr_sum / nbr);
	sheet.getRange('AP26').setValue(irr_sum);
	sheet.getRange('AP27').setValue(nbr);
    }
}

function average_irr() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('IRR');
    if (sheet) {
	sheet.activate();
    }
    else {
	return;
    }

    average_5y_irr(sheet);
    average_10y_irr(sheet);
    average_15y_irr(sheet);
}

function calculate_irr() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheetname = 'IRR';
    IRR = spreadsheet.getSheetByName(sheetname);
    if (IRR) {
	//IRR.clear();
	IRR.activate();
    }
    else {
	IRR = spreadsheet.insertSheet(sheetname, spreadsheet.getNumSheets());
    }
    
    var label_range = IRR.getRange('F1:J1');
    label_range.setFontWeight('bold');
    var labels = [
	['Price', 'Current return', 'IRR', 'Stay in Market (Y)', '']
    ];
    label_range.setValues(labels);

    var ticket = IRR.getRange('A1:A1');
    var start_of_year = 2024;
    var end_of_year = 2000;
    var year_nbr = start_of_year - end_of_year + 1;
    var year_range = IRR.getRange('A4:J60');
    var month_range = IRR.getRange('B4:K60');
    var day_range = IRR.getRange('C4:L60');
    var date_range = IRR.getRange('D4:M60');
    var pricefunc_range = IRR.getRange('E4:N60');
    var r = 1;
    var ycell = year_range.getCell(r, 1);
    var mcell = month_range.getCell(r,1);
    var dcell = day_range.getCell(r,1);
    var datecell = date_range.getCell(r,1);
    var pricefuncell = pricefunc_range.getCell(r,1);

    for (var y = start_of_year; y >= end_of_year; --y) {
	ycell.setValue(y);
	ycell.setBackground('#ddeeee');
	mcell.setValue(1);
	// dcell.setValue(1);
	datecell.setFormula('=date('+ ycell.getA1Notation() + ',' + mcell.getA1Notation() + ',' + dcell.getA1Notation() + ')');
	r += 2;
	pricefuncell.setFormula('=googlefinance(\"' + ticket.getCell(1,1).getValue() + '\", \"price\",' + datecell.getA1Notation() + ')');
	ycell = year_range.getCell(r,1);
	mcell = month_range.getCell(r,1);
	dcell = day_range.getCell(r,1);
	datecell = date_range.getCell(r,1);
	pricefuncell = pricefunc_range.getCell(r,1);
    }

    var present_price = IRR.getRange('F3:F3');
    present_price.getCell(1,1).setFormula('=googlefinance(\"' + ticket.getCell(1,1).getValue() + '\")');
    // Calculate current return of the year
    yy = 0;
    r = 1;
    var price_range = IRR.getRange('F3:F300');
    var ct_range = IRR.getRange('G3:G300');
    var pricecell = price_range.getCell(r,1);

    for (var ctcell = ct_range.getCell(r,1); yy < year_nbr;
	 r += 2, ++yy,
	 ctcell = ct_range.getCell(r,1),
	 pricecell = price_range.getCell(r,1)) 
    {
	//var y0 = pricecell.getValue();
	//var y1 = price_range.getCell(r+2,1).getValue();
	//roicell.setValue(y1/(y1-y0));
	ctcell.setFormula('=' + '(' + pricecell.getA1Notation() + '-' + price_range.getCell(r + 2, 1).getA1Notation() + ')/' + price_range.getCell(r + 2, 1).getA1Notation());
	ctcell.setNumberFormat("##.##%");
    }
    calculate_irr_values(ticket, year_nbr, price_range);
    // calculate_roi(ticket, year_nbr, price_range);
    // calculate_roi_array(ticket, year_nbr, price_range);
    SpreadsheetApp.flush();
}

function calculate_irr_values(ticket, year_nbr, price_range)
{
    // Calculate IRR in different starting years
    var irr_range = IRR.getRange('H5:AY300');
    var yylabel_range = IRR.getRange('I5:AY300');
    var yylabel_c_nbr = yylabel_range.getNumColumns();

    for (var c = 1; c <= yylabel_c_nbr; c += 2)
    {
	var r = c;
	var yy = 1;
	var current_price = price_range.getCell(r,1);
	var yylabel_cell = yylabel_range.getCell(r,c);

	for (var irrcell = irr_range.getCell(r,c); yy <= year_nbr; 
	     yy++, r += 2,
	     irrcell = irr_range.getCell(r,c),
	     yylabel_cell = yylabel_range.getCell(r,c)
	    ) {
	    // power(y1/y0, 1/yy) - 1
	    irrcell.setFormula(
		'=power(' + current_price.getA1Notation() + '/' + price_range.getCell(r + 2, 1).getA1Notation() + ',1/' + yy + ') - 1'
	    );
	    irrcell.setNumberFormat("##0.##%");
	    yylabel_cell.setValue(yy);
	    yylabel_cell.setHorizontalAlignment('left');
	    yylabel_cell.setFontWeight('bold');
	    yylabel_cell.setFontColor('blue');
	}
    }
}

function calculate_roi(ticket, year_nbr, price_range)
{
    // Calculate IRR, ROI
    r = 1;
    yy = 1;
    var present_price = IRR.getRange('F3:F3');
    present_price.getCell(1,1).setFormula('=googlefinance(\"' + ticket.getCell(1,1).getValue() + '\")');
    var irr_range = IRR.getRange('H5:H60');
    var roi_range = IRR.getRange('J5:J60');
    var yylabel_range = IRR.getRange('I5:I60');
    var yylabel_cell = yylabel_range.getCell(r,1);
    var roi_cell = roi_range.getCell(r,1);
    for (var irrcell = irr_range.getCell(r,1); yy <= year_nbr; 
	 yy++, r += 2,
	 irrcell = irr_range.getCell(r,1),
	 yylabel_cell = yylabel_range.getCell(r,1),
	 roi_cell = roi_range.getCell(r,1)
	) {
	// power(y1/y0, 1/yy) - 1
	irrcell.setFormula(
	    '=power(' + present_price.getCell(1,1).getA1Notation() + '/' + price_range.getCell(r + 2, 1).getA1Notation() + ',1/' + yy + ') - 1'
	);
	irrcell.setNumberFormat("##.##%");
	yylabel_cell.setValue(yy);
	roi_cell.setFormula('=(' + present_price.getCell(1,1).getA1Notation() + '-' + price_range.getCell(r + 2,1).getA1Notation() + ')' + '/' + price_range.getCell(r + 2,1).getA1Notation());
	roi_cell.setNumberFormat("##.##%");
    }
}

/**
 * A custom function that converts meters to miles.
 *
 * @param {Number} meters The distance in meters.
 * @return {Number} The distance in miles.
 */
function metersToMiles(meters) {
    if (typeof meters != 'number') {
	return null;
    }
    return meters / 1000 * 0.621371;
}

/**
 * A custom function that gets the driving distance between two addresses.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Number} The distance in meters.
 */
function drivingDistance(origin, destination) {
    var directions = getDirections_(origin, destination);
    return directions.routes[0].legs[0].distance.value;
}

/**
 * A function that adds headers and some initial data to the spreadsheet.
 */
function prepareSheet_() {
    var sheet = SpreadsheetApp.getActiveSheet().setName('Settings');
    var headers = [
	'Start Address',
	'End Address',
	'Driving Distance (meters)',
	'Driving Distance (miles)'];
    var initialData = [
	'350 5th Ave, New York, NY 10118',
	'405 Lexington Ave, New York, NY 10174'];
    sheet.getRange('A1:D1').setValues([headers]).setFontWeight('bold');
    sheet.getRange('A2:B2').setValues([initialData]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 4);
}

/**
 * Creates a new sheet containing step-by-step directions between the two
 * addresses on the "Settings" sheet that the user selected.
 */
function generateStepByStep_() {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName('Settings');
    settingsSheet.activate();

    // Prompt the user for a row number.
    var selectedRow = Browser.inputBox('Generate step-by-step',
				       'Please enter the row number of the addresses to use' +
				       ' (for example, "2"):',
				       Browser.Buttons.OK_CANCEL);
    if (selectedRow == 'cancel') {
	return;
    }
    var rowNumber = Number(selectedRow);
    if (isNaN(rowNumber) || rowNumber < 2 ||
	rowNumber > settingsSheet.getLastRow()) {
	Browser.msgBox('Error',
		       Utilities.formatString('Row "%s" is not valid.', selectedRow),
		       Browser.Buttons.OK);
	return;
    }

    // Retrieve the addresses in that row.
    var row = settingsSheet.getRange(rowNumber, 1, 1, 2);
    var rowValues = row.getValues();
    var origin = rowValues[0][0];
    var destination = rowValues[0][1];
    if (!origin || !destination) {
	Browser.msgBox('Error', 'Row does not contain two addresses.',
		       Browser.Buttons.OK);
	return;
    }

    // Get the raw directions information.
    var directions = getDirections_(origin, destination);

    // Create a new sheet and append the steps in the directions.
    var sheetName = 'Driving Directions for Row ' + rowNumber;
    var directionsSheet = spreadsheet.getSheetByName(sheetName);
    if (directionsSheet) {
	directionsSheet.clear();
	directionsSheet.activate();
    } else {
	directionsSheet =
	    spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
    }
    var sheetTitle = Utilities.formatString('Driving Directions from %s to %s',
					    origin, destination);
    var headers = [
	[sheetTitle, '', ''],
	['Step', 'Distance (Meters)', 'Distance (Miles)']
    ];
    var newRows = [];
    for (var i = 0; i < directions.routes[0].legs[0].steps.length; i++) {
	var step = directions.routes[0].legs[0].steps[i];
	// Remove HTML tags from the instructions.
	var instructions = step.html_instructions.replace(/<br>|<div.*?>/g, '\n')
	    .replace(/<.*?>/g, '');
	newRows.push([
	    instructions,
	    step.distance.value
	]);
    }
    directionsSheet.getRange(1, 1, headers.length, 3).setValues(headers);
    directionsSheet.getRange(headers.length + 1, 1, newRows.length, 2)
	.setValues(newRows);
    directionsSheet.getRange(headers.length + 1, 3, newRows.length, 1)
	.setFormulaR1C1('=METERSTOMILES(R[0]C[-1])');

    // Format the new sheet.
    directionsSheet.getRange('A1:C1').merge().setBackground('#ddddee');
    directionsSheet.getRange('A1:2').setFontWeight('bold');
    directionsSheet.setColumnWidth(1, 500);
    directionsSheet.getRange('B2:C').setVerticalAlignment('top');
    directionsSheet.getRange('C2:C').setNumberFormat('0.00');
    var stepsRange = directionsSheet.getDataRange()
	.offset(2, 0, directionsSheet.getLastRow() - 2);
    setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
    directionsSheet.setFrozenRows(2);
    SpreadsheetApp.flush();
}

/**
 * Sets the background colors for alternating rows within the range.
 * @param {Range} range The range to change the background colors of.
 * @param {string} oddColor The color to apply to odd rows (relative to the
 *     start of the range).
 * @param {string} evenColor The color to apply to even rows (relative to the
 *     start of the range).
 */
function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
    var backgrounds = [];
    for (var row = 1; row <= range.getNumRows(); row++) {
	var rowBackgrounds = [];
	for (var column = 1; column <= range.getNumColumns(); column++) {
	    if (row % 2 == 0) {
		rowBackgrounds.push(evenColor);
	    } else {
		rowBackgrounds.push(oddColor);
	    }
	}
	backgrounds.push(rowBackgrounds);
    }
    range.setBackgrounds(backgrounds);
}

/**
 * A shared helper function used to obtain the full set of directions
 * information between two addresses. Uses the Apps Script Maps Service.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Object} The directions response object.
 */
function getDirections_(origin, destination) {
    var directionFinder = Maps.newDirectionFinder();
    directionFinder.setOrigin(origin);
    directionFinder.setDestination(destination);
    var directions = directionFinder.getDirections();
    if (directions.status !== 'OK') {
	throw directions.error_message;
    }
    return directions;
}
