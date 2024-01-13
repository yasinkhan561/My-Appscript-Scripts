function fetchStockData(symbol) {
	var apiKey = "60gHhAkGGpYsTriA0pBCoV2xX1yUDEVB";
	var url =
		"https://api.polygon.io/v3/reference/tickers/" +
		symbol +
		"?apiKey=" +
		apiKey;
	var options = {
		muteHttpExceptions: true,
	};
	var response = UrlFetchApp.fetch(url, options);
	var json = response.getContentText();
	try {
		var data = JSON.parse(json);
		Logger.log(data);
	} catch (e) {
		Logger.log("JSON parse error: " + e);
	}
}

function appendDividendData() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getActiveSheet();
	var apiKey = "60gHhAkGGpYsTriA0pBCoV2xX1yUDEVB";
	var url = "https://api.polygon.io/v3/reference/dividends?apiKey=" + apiKey;
	var options = {
		muteHttpExceptions: true,
	};
	var response = UrlFetchApp.fetch(url, options);
	var json = response.getContentText();
	var data = JSON.parse(json);
	var results = data.results;

	for (i in results) {
		var row = results[i];
		Logger.log(row);
		var newRow = [row.ticker, row.cash_amount, row.pay_date, row.period];
		Logger.log(newRow);
	}

	// sheet.appendRow([data.ticker, data.amount, data.payDate, data.period]);
}
