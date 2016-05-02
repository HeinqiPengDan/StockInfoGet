/// <reference path="../App.js" />

(function () {
	"use strict";

	var initCheckBox = function () {

		$("#column_show_all").prop("checked", true);
		$("#column_show_all")[0].parentElement.className = "checked_label";

		$("input[name='column_show']:enabled").prop("checked", $("#column_show_all").prop("checked"));
		//change all the other checkbox parent label color
		$("input[name='column_show']").each(function () {
			if (this.checked) {
				this.parentElement.className = "checked_label";
			} else {
				this.parentElement.className = "";
			}
		});
	};

	var column_amount = 7;

	var allCheckedOrNot = function () {
		var checkedCount = $("input[name='column_show']:checked").length;

		//If time column is disabled, then we have less than one columns
		if ($("#column_show_time").prop("disabled") == true) {
			column_amount = 6;
		} else {
			column_amount = 7;
		}

		if (checkedCount == column_amount) { //All columns' checkbox are selected
			$("#column_show_all").prop("checked", true);
			$("#column_show_all")[0].parentElement.className = "checked_label";

		} else {
			$("#column_show_all").prop("checked", false);
			$("#column_show_all")[0].parentElement.className = "";

		}

		if (event.srcElement.checked) {
			event.srcElement.parentElement.className = "checked_label";
		} else {
			event.srcElement.parentElement.className = "";
		}
	};

	//Column label hover function
	var labelOn = function () {
		if (event.srcElement.className != "checked_label" && event.srcElement.className != "disabled_label") { //Only work if checkbox is not checked and enabled
			event.srcElement.className = "selected_label";
		}
	}

	//Column label out function
	var labelOut = function () {
		if (event.srcElement.className != "checked_label" && event.srcElement.className != "disabled_label") { //Only work if checkbox is not checked and enabled
			event.srcElement.className = "";
		}
	}

	//Disable column_show_time checkbox
	var disableTimeCheckbox = function () {
		var market = $("#market").val();
		if (market == "Hong Kong(HK)" || market == "United States(US)") {
			$("#column_show_time").prop("checked", false);
			$("#column_show_time").prop("disabled", true);
			$("#column_show_time")[0].parentElement.className = "disabled_label";

			//If time column is disable, but all others are checked, we also need "All" is checked
			var checkedCount = $("input[name='column_show']:checked").length;
			if (checkedCount == 6) { //All other columns' checkbox are selected
				$("#column_show_all").prop("checked", true);
				$("#column_show_all")[0].parentElement.className = "checked_label";

			} else {
				$("#column_show_all").prop("checked", false);
				$("#column_show_all")[0].parentElement.className = "";
			}

		} else { //Chinese market
			$("#column_show_time").prop("disabled", false);
			$("#column_show_time")[0].parentElement.className = "";

			//if all is select, then reselect this check box
			if ($("#column_show_all").prop("checked") == true) {
				$("#column_show_time").prop("checked", true);
				$("#column_show_time")[0].parentElement.className = "checked_label";
			}

		}
	}

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();
			initCheckBox();

			$('#generate').click(showData);

			$('#column_show_all').click(checkAndUnCheckAll);

			$("input[name='column_show']").on("click", allCheckedOrNot);

			$("label[name='lbl_column']").on("mouseover", labelOn);

			$("label[name='lbl_column']").on("mouseout", labelOut);

			$("#market").on("change", disableTimeCheckbox);

			$("#stock_id").keypress(function (event) {
				if (event.keyCode == 13) {
					$('#generate').click();
				}
			});

		});
	};

	// Reads data from current document selection and displays a notification
	function getDataFromSelection() {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				app.showNotification('The selected text is:', '"' + result.value + '"');
			} else {
				app.showNotification('Error:', result.error.message);
			}
		});
	}

	// Check or uncheck all others checkboxes once "All" is clicked
	function checkAndUnCheckAll() {
		$("input[name='column_show']:enabled").prop("checked", $("#column_show_all").prop("checked"));

		//Change "all" checkbox parent label color
		if (event.srcElement.checked) {
			event.srcElement.parentElement.className = "checked_label";
		} else {
			event.srcElement.parentElement.className = "";
		}
		//change all the other checkbox parent label color
		$("input[name='column_show']:enabled").each(function () {
			if (this.checked) {
				this.parentElement.className = "checked_label";
			} else {
				this.parentElement.className = "";
			}
		});
	}

})();

var stockColumnHeaders = ['Name', 'Code', 'Date', 'Time', 'OpenningPrice', 'ClosingPrice', 'CurrentPrice'];

//Function for button click
function showData() {
	//Check stock id input
	var stockId = $("#stock_id").val();

	if (stockId == "") {
		showMarketData();
	} else {
		//If user select no columns to display, just return
		var checkedCount = $("input[name='column_show']:checked").length;
		if (checkedCount == 0) {
			app.showNotification("Please select display columns!");
			return;
		}
		showStockData();
	}
}

//retrieve and show the market data
function showMarketData(symbol) {

	var market = $("#market").val();

	//var marketColumns = ['Name', 'Date', 'Current Dot', 'Rate', 'Growth', 'Start Dot', 'Close Dot', 'High Dot', 'Low Dot', 'Turnover'];
	var marketColumns = new Array();
	var xmlhttp;
	var i;

	marketColumns[0] = "Name";
	marketColumns[1] = "Code";
	marketColumns[2] = "Date";
	marketColumns[3] = "Current Dot";
	marketColumns[4] = "Rate";
	marketColumns[5] = "Growth";
	if (market == "United States(US)" || market == "Hong Kong(HK)") {
		marketColumns[6] = "Start Dot";
		marketColumns[7] = "Close Dot";
		marketColumns[8] = "High Dot";
		marketColumns[9] = "Low Dot";

	}

	if (window.XMLHttpRequest) { // code for IE7+, Firefox, Chrome, Opera, Safari
		xmlhttp = new XMLHttpRequest();
	} else { // code for IE6, IE5
		xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	}

	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 3 && xmlhttp.status == 200) {

			var xmlDoc = JSON.parse(xmlhttp.responseText);

			var index = 0;
			var marketData = new Array();
			var dateValue;
			var cnFlag = 0;
			var marketCode;
			for (var eachMarket in xmlDoc.retData.market) {

				//don't show this always
				if (eachMarket == "INX") {
					continue;
				}

				//Show market data based on different selection
				if ((eachMarket == "shanghai" || eachMarket == "shenzhen") && market != "Chinese Mainland(CN)") {
					continue;
				}

				if ((eachMarket == "DJI" || eachMarket == "IXIC") && market != "United States(US)") {
					continue;
				}

				if (eachMarket == "HSI" && market != "Hong Kong(HK)") {
					dateValue = xmlDoc.retData.market[eachMarket].date;
					continue;
				}

				if (eachMarket == "shanghai") {
					marketCode = "SH000001"
				}
				if (eachMarket == "shenzhen") {
					marketCode = "SZ399001"
				}
				if (eachMarket == "DJI") {
					marketCode = "DJIA"
				}
				if (eachMarket == "IXIC") {
					marketCode = "NASDAQ"
				}
				if (eachMarket == "HSI") {
					marketCode = "HSI"
				}

				if (market == "Chinese Mainland(CN)") {
					marketData[index] = new Array();
					marketData[index][0] = xmlDoc.retData.market[eachMarket].name;
					marketData[index][1] = marketCode;
					marketData[index][3] = xmlDoc.retData.market[eachMarket].curdot;
					marketData[index][4] = xmlDoc.retData.market[eachMarket].rate;
					marketData[index][5] = xmlDoc.retData.market[eachMarket].curprice;
					cnFlag = 1;
				} else {

					marketData[index] = new Array();
					marketData[index][0] = xmlDoc.retData.market[eachMarket].name;
					marketData[index][1] = marketCode;
					marketData[index][2] = xmlDoc.retData.market[eachMarket].date;
					marketData[index][3] = xmlDoc.retData.market[eachMarket].curdot;
					marketData[index][4] = xmlDoc.retData.market[eachMarket].rate;
					marketData[index][5] = xmlDoc.retData.market[eachMarket].growth;
					marketData[index][6] = xmlDoc.retData.market[eachMarket].startdot;
					marketData[index][7] = xmlDoc.retData.market[eachMarket].closedot;
					marketData[index][8] = xmlDoc.retData.market[eachMarket].hdot;
					marketData[index][9] = xmlDoc.retData.market[eachMarket].ldot;

				}
				index++;

			}

			if (cnFlag == 1) {
				for (i = 0; i < index; i++) {
					marketData[i][2] = dateValue;
				}
			}

			var marketTableData = new Office.TableData(marketData, marketColumns);

			Office.context.document.setSelectedDataAsync(
				marketTableData, {
				coercionType : Office.BindingType.Table
			},
				function (asyncResult) {
				if (asyncResult.status == "failed") {
					app.showNotification("Error when getting market info: " + asyncResult.error.message);
				} else {
					app.showNotification("Stock info is retrieved successfully!");
				}
			});
		}
	}

	var url = "https://apis.baidu.com/apistore/stockservice/hkstock?stockid=00168&list=1";
	xmlhttp.open("GET", url, true);
	xmlhttp.setRequestHeader("apikey", "1d305e127b69ec773b98d0cea3ace133");
	xmlhttp.send();

}

//retrieve and show the stock data
function showStockData(symbol) {
	var xmlhttp;
	var i;

	//Filter display column based on user selected.
	var filterStockColumnHeaders = new Array();
	var columns = new Array();
	var columnsIndex = 0;

	$("input[name='column_show']").each(function () {
		if (this.checked) {
			columns[columnsIndex] = true;
		} else {
			columns[columnsIndex] = false;
		}
		columnsIndex = columnsIndex + 1;
	});

	var filterStockColumnHeadersIndex = 0;
	for (var k = 0; k < stockColumnHeaders.length; k++) {
		if (columns[k] == true) {
			filterStockColumnHeaders[filterStockColumnHeadersIndex] = stockColumnHeaders[k];
			filterStockColumnHeadersIndex++;
		}
	}

	if (window.XMLHttpRequest) { // code for IE7+, Firefox, Chrome, Opera, Safari
		xmlhttp = new XMLHttpRequest();
	} else { // code for IE6, IE5
		xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	}

	xmlhttp.onreadystatechange = function () {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

			var xmlDoc = JSON.parse(xmlhttp.responseText);
			//    document.getElementById("myDiv").innerHTML=xmlDoc.;
			//			alert(xmlhttp.responseText);
			var stockData = new Array();
			for (i = 0; i < xmlDoc.retData.stockinfo.length; i++) {
				stockData[i] = new Array();
				var stockColIndex = 0;
				if (columns[0] == true) {
					var name = processNull(xmlDoc.retData.stockinfo[i].name);
					if (name == "" || name == "FAILED") {
						stockData[i][stockColIndex] = "No response data with stock id " + processNull(xmlDoc.retData.stockinfo[i].code) + ", invalid id?";
					} else {
						stockData[i][stockColIndex] = name;
					}
					stockColIndex++;
				}

				if (columns[1] == true) {
					stockData[i][stockColIndex] = processStockCode(xmlDoc.retData.stockinfo[i].code);
					stockColIndex++;
				}

				if (columns[2] == true) {
					stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].date);
					stockColIndex++;
				}

				if (columns[3] == true) {
					stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].time);
					stockColIndex++;
				}

				if (columns[4] == true) {
					stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].openningPrice);
					if (stockData[i][stockColIndex] == "")
						stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].OpenningPrice);
					stockColIndex++;
				}

				if (columns[5] == true) {
					stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].closingPrice);
					stockColIndex++;
				}

				if (columns[6] == true) {
					stockData[i][stockColIndex] = processNull(xmlDoc.retData.stockinfo[i].currentPrice);
					stockColIndex++;
				}
			}

			//            var stockTableData = new Office.TableData(stockData, filterStockColumnHeaders);
			var stockTableData = new Office.TableData()
				stockTableData.headers = filterStockColumnHeaders;
			stockTableData.rows = stockData;
			var beginFlag = 0;
			//var stockColumnHeaders = ['Name', 'Code', 'Date', 'Time', 'OpenningPrice', 'ClosingPrice', 'CurrentPrice'];
			var FormatString = " ";
			for (i = 0; i < filterStockColumnHeaders.length; i++) {
				if (filterStockColumnHeaders[i] != "Name" && filterStockColumnHeaders[i] != "Code") {
					if (beginFlag != 0) {
						FormatString = FormatString + ",";
					} else {
						beginFlag = 1;
					}
					if (filterStockColumnHeaders[i] == "Date") {
						if (market == "Chinese Mainland(CN)") {
							FormatString = FormatString + "{cells: {row:" + String(i) + "}, format: {numberFormat:" + "\"m/d/yyyy\"" + "}}";
						} else {
							FormatString = FormatString + "{cells: {row:" + String(i) + "}, format: {numberFormat:" + "\"m/d/yyyy h:mm:ss\"" + "}}";
						}
					}
					if (filterStockColumnHeaders[i] == "Time") {
						FormatString = FormatString + "{cells: {row:" + String(i) + "}, format: {numberFormat:" + "\"h:mm:ss\"" + "}}";
					}
					if (filterStockColumnHeaders[i] == "OpenningPrice" || filterStockColumnHeaders[i] == "ClosingPrice" || filterStockColumnHeaders[i] == "CurrentPrice") {
						FormatString = FormatString + "{cells: {row:" + String(i) + "}, format: {numberFormat:" + "\"####.#\"" + "}}";
					}
				}

			}
			if (beginFlag != 0) {
				FormatString = "[" + FormatString + "]";
			}
			Office.context.document.setSelectedDataAsync(
				stockTableData, [{
						coercionType : Office.BindingType.Table
					}, {
						cellFormat : FormatString
					}
				],
				function (asyncResult) {
				if (asyncResult.status == "failed") {
					app.showNotification("Error when getting stock info: " + asyncResult.error.message);
				} else {
					app.showNotification("Stock info is retrieved successfully!");
				}
			});
		}
	}

	var url = "https://apis.baidu.com/apistore/stockservice/stock?list=1";
	//check market input
	var market = $("#market").val();
	if (market == "Chinese Mainland(CN)") { //Chinese market
		url = "https://apis.baidu.com/apistore/stockservice/stock?list=1";
	} else if (market == "Hong Kong(HK)") { //HK market
		url = " https://apis.baidu.com/apistore/stockservice/hkstock?list=1";
	} else { //US market
		url = " https://apis.baidu.com/apistore/stockservice/usastock?list=1";
	}

	//Check stock id input
	var stockId = $("#stock_id").val();
	url = url + "&stockid=" + stockId.toLowerCase();

	xmlhttp.open("GET", url, true);
	xmlhttp.setRequestHeader("apikey", "1d305e127b69ec773b98d0cea3ace133");
	xmlhttp.send();

}

//Get Json Length
function getJsonLength(jsonData) {

	var jsonLength = 0;

	for (var item in jsonData) {

		jsonLength++;

	}

	return jsonLength;

}

//Process null or undefined
function processNull(obj) {
	if (obj == null || typeof(obj) == "undefined") {
		return "";
	} else {
		return obj;
	}
}

//Process stock code (stock ID), 00700 will display as 700 in excel,
//We need handle this
function processStockCode(code) {
	if (code == null || typeof(code) == "undefined") {
		return "";
	} else {
		if (!isNaN(code) && code.length > 1 && code.substring(0, 1) == "0") { //code is a pure numberic string and start with 0
			return "'" + code;
		} else {
			return code
		}
	}
}