/**
 *
 */

var GainLoss2026_sht;
var GainLoss2026_Summary_sht;
var TickerSheets = [];
const Data_source_sheet = 'Gain-Loss.US.2026.Log';
const Data_output_sheet = 'Gain-Loss.US.2026';

function calculate_gainloss() {
    var spreadsheet = SpreadsheetApp.getActive();
    var gainlossSheet = spreadsheet.getSheetByName(Data_source_sheet);

    if (gainlossSheet) {
	GainLoss2026_sht = gainlossSheet;
	// gainlossSheet.activate();
	create_summary_sheet(Data_output_sheet);
    }
    else {
	message_box("No '" + sheetname + "' alive");
	return;
    }

    var tickers = [
	"DRAM",
	"MU",
	"ADBE",
	"ADI",
	"AMD",
	"ASTS",
	"BRKB",
	"CBRS",
	"CLS",
	"COHR",
	"CRWD",
	"IBM",
	"LITE",
	"MRVL",
	"MSFT",
	"NBIS",
	"NKE",
	"NLR",
	"ON",
	"ORCL",
	"PLTR",
	"QQQM",
	"SKHY",
	"SMCI",
	"SNDK",
	"SPCX",
	"T",
	"TMUS",
	"TSLA",
	"TXN",
	"VOX",
	"QLD",
	"VGT",
	"XLK",
	"MSTR",
	"IBIT",
	"IEI"
    ];

    var tickers_in_eq = [
	"BATS:DRAM",
	"MU",
	"ADBE",
	"ADI",
	"AMD",
	"ASTS",
	"BRK.B",
	"CBRS",
	"CLS",
	"COHR",
	"CRWD",
	"IBM",
	"LITE",
	"MRVL",
	"MSFT",
	"NBIS",
	"NKE",
	"NLR",
	"ON",
	"ORCL",
	"PLTR",
	"QQQM",
	"SKHY",
	"SMCI",
	"SNDK",
	"SPCX",
	"T",
	"TMUS",
	"TSLA",
	"TXN",
	"VOX",	
	"QLD",
	"VGT",
	"XLK",
	"MSTR",
	"IBIT",
	"IEI"
    ];

    for (var i = 0; i < tickers.length; ++i) {
	create_ticker_sheet(tickers[i]);
	create_ticker_data(tickers[i], tickers_in_eq[i], i); // Array index in TickerSheets
    }
}

function create_summary_sheet(name) {
    var spreadsheet = SpreadsheetApp.getActive();
    var sht = spreadsheet.getSheetByName(Data_output_sheet);

    if (sht) {
	sht.clear();
    }
    else {
	sht = spreadsheet.insertSheet(name, spreadsheet.getNumSheets());
	sht.getRange("C2:J").setNumberFormat("#,##0.00");
    }
    GainLoss2026_Summary_sht = sht;
    GainLoss2026_Summary_sht.appendRow(["", "Ticker", "Price", "Unit cost", "", "Total Cost", "Sum of Share", "Sum of Realized", "Market Value", "Un-realized ($)", "Un-realized (%)", "Cumulative cost", "Spent cost", "Realized gain (%)"]);
}

function create_ticker_sheet(name) {
    var sheetname = name + '.2026';
    var spreadsheet = SpreadsheetApp.getActive();
    var sht = spreadsheet.getSheetByName(sheetname);

    if (sht) {
	sht.clear();
    }
    else {
	sht = spreadsheet.insertSheet(sheetname, spreadsheet.getNumSheets());
    }
    TickerSheets.push(sht);
}

function collet_ticker_entries(ticker, ticker_in_eq) {
    var rowNumber = GainLoss2026_sht.getMaxRows();
    var range = GainLoss2026_sht.getRange("A1:M" + rowNumber);
    SpreadsheetApp.getActive().setNamedRange("Symbol", range);
    var values = range.getValues();

    var share = 0, addup = 0, realized = 0, unitcost = 0, sum_realized = 0, sum_buy = 0;
    var data = [];
    data.push(["Date", "Action", "Price", "Unit cost", "Quantity", "Total cost", "Sum of share", "Sum of cost", "Realized", "Sum of Realized"]);

    values.forEach(function (row) {
	var sym = row[0];
	if (sym == ticker) {
	    var qty = row[1];
	    var price = row[2];
	    var action = row[3];
	    var descrip = row[4];
	    var date = row[5];
	    var money = row[8];

	    // BUY
	    if (qty > 0) {
		share += qty;
		addup += money;
		sum_buy += money;
		if (share == 0) {
		    sum_realized += addup;
		    realized = addup;
		    addup = 0;
		    unitcost = 0;
		}
		else {
		    unitcost = Math.abs(addup / share);
		}
		// Logger.log("BUY>   Price: " + price + " | unit cost: " + unitcost + " | quantity: " + qty + " | Realized: " + realized + " | Total realized: " + sum_realized + " | Total share/cost: " + share + '/' + addup);
		if (share != 0) {
		    data.push([date, action, price, unitcost, qty, money, share, addup]);
		}
		else {
		    data.push([date, action, price, unitcost, qty, money, share, addup, realized, sum_realized]);
		}
	    }
	    // SELL
	    else if (qty < 0) {
		share += qty;
		addup += money;
		realized = (price - unitcost) * -qty; // -qty, make qty positive
		if (share == 0) {
		    addup = 0; // Reset floating errors
		}
		else if (share < 0) {
		    // data log order is incorrect
		    addup -= realized;
		}
		else {
		    addup -= realized;
		}
		sum_realized += realized;
		// Logger.log("SELL>   Price: " + price + " | unit cost: " + unitcost + " | quantity: " + qty + " | Realized: " + realized + " | Total realized: " + sum_realized + " | Cost add up: " + addup);
		data.push([date, action, price, unitcost, qty, money, share, addup, realized, sum_realized]);
		// Logger.log("	     Total realized: " + sum_realized);
	    }
	    // Dividend
	    else if (qty == 0) {
		if (share == 0) {}
		else {
		    addup += money;
		}
		realized = money;
		sum_realized += money;
		data.push([date, action, price, unitcost, qty, money, share, addup, realized, sum_realized]);
	    }
	}
    });

    // Logger.log(ticker + " Realized: " + sum_realized);
    data.push(["", "Ticker", "Price", "Unit cost", "", "Total Cost", "Sum of Share", "Sum of Realized", "Market Value", "Un-realized ($)", "Un-realized (%)", "Cumulative cost", "Spent cost", "Realized gain (%)"]);

    {
	var price_eq = "=googlefinance(\"" + ticker_in_eq + "\")";
	// Gen equation: market value
	var price = "indirect(\"C\" & row())";
	var shr = "indirect(\"G\" & row())"
	var mktval_eq = "=if (" + shr +", " + price + '*' + shr + ',' + "\"\")";

	// Gen equation: un-realized value
	var ref_mktval = "indirect(\"I\" & row())";
	var ref_total_cost = "indirect(\"F\" & row())"
	var unrealized_eq = '=' + ref_mktval + '+' + ref_total_cost;

	// Un-realized gain (%)
	var ref_unrealized = "indirect(\"J\" & row())"
	var unrealized_gain_pcnt_eq = "=if (" + ref_total_cost + ", " + ref_unrealized + " / -" + ref_total_cost + ", \"\")";

	// Spend cost ($)
	var ref_sum_buy = "indirect(\"L\" & row())";
	var spend_cost_eq = "=-" + ref_sum_buy + '+' + ref_total_cost;

	// Realized gain (%)
	var ref_realized = "indirect(\"H\" & row())";
	var ref_spend_cost = "indirect(\"M\" & row())";
	var realized_gain_pcnt_eq = '=' + ref_realized + '/' + ref_spend_cost;

	if (addup == 0) addup = "";
	if (share == 0) share = "";
	data.push(["", ticker, price_eq, unitcost, "", addup, share, sum_realized, mktval_eq, unrealized_eq, unrealized_gain_pcnt_eq, sum_buy, spend_cost_eq, realized_gain_pcnt_eq]);
    }
    return data;

    // message_box("Share: " + share + " sum: " + addup);

    // values.forEach(function (row) {
    //	 row.forEach(function (col) {
    //	   Logger.log(col);
    //	 });
    // });

    // message_box("Row number: " + rowNumber + " / Column number: " + colNumber);
}

function create_ticker_data(ticker, ticker_eq, sht_id) {
    TickerSheets[sht_id].getRange("C2:J").setNumberFormat("#,##0.00");
    var entries = collet_ticker_entries(ticker, ticker_eq);
    for (var i = 0; i < entries.length; ++i) {
	TickerSheets[sht_id].appendRow(entries[i]);
    }
    GainLoss2026_Summary_sht.appendRow(entries[entries.length - 1]);
}

function message_box(msg) {
    SpreadsheetApp.getUi().alert(msg);
}
