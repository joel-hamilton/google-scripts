function updateSelectedStocks() {
  updateStocks(true);
}

function updateStocks(selectedOnly = false) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeRangeList = sheet.getActiveRangeList();
  var tickers;

  if (activeRangeList && selectedOnly) {
    // only updating selected lines
    let ranges = activeRangeList.getRanges();
    ranges.forEach((range) => {
      var rowNum = range.getRow();
      var ticker = sheet.getRange(`A${rowNum}`).getValue();
      let info = getStockInfo(ticker);
      setRowValues(sheet, rowNum, info);
    });
  } else {
    // batch update everything
    setHeaders(sheet);
    var startRow = 2;
    var tickers = sheet.getRange(`A${startRow}:D500`).getValues();

    tickers
      .filter((arr) => arr.length)
      .map((item) => item[0])
      .forEach((ticker, index) => {
        let rowNum = index + startRow;
        let info = getStockInfo(ticker);
        setRowValues(sheet, rowNum, info);
      });
  }
}

function setHeaders(sheet, headers) {
  var sampleInfo = getStockInfo("AAPL");
  var headers = Object.keys(sampleInfo);
  headers.unshift("Last Updated");
  setRowValues(sheet, 1, headers);
}

function getHeaders(sheet) {
  return sheet.getRange(`A1:ZZ1`).getValues()[0];
}

function setRowValues(sheet, rowNum, info) {
  let headers = getHeaders(sheet);
  headers.forEach((header, index) => {
    if (Object.keys(info).includes(header)) {
      var cell = numToSSColumn(index + 1) + rowNum;
      sheet.getRange(cell).setValue(info[header]);
    }
  });
}

function getStockInfo(ticker) {
  var info = {};
  var url = "https://finviz.com/quote.ashx?t=" + ticker;

  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var html = res.getContentText();
  var $ = Cheerio.load(html);

  $(".snapshot-table2 .snapshot-td2-cp").each(function () {
    var field = $(this).text();
    var value = $(this).next().text();
    info[field] = parseValue(value);
  });

  return info;
}

function parseValue(rawValue) {
  // if value ends in M or B, convert to millions or billions
  if (rawValue.substring(rawValue.length - 1) === "M") {
    return parseFloat(rawValue.substring(0, rawValue.length - 1)) * 1000 * 1000;
  }

  if (rawValue.substring(rawValue.length - 1) === "B") {
    return (
      parseFloat(rawValue.substring(0, rawValue.length - 1)) *
      1000 *
      1000 *
      1000
    );
  }

  return rawValue;
}

function numToSSColumn(num) {
  var s = "",
    t;

  while (num > 0) {
    t = (num - 1) % 26;
    s = String.fromCharCode(65 + t) + s;
    num = ((num - t) / 26) | 0;
  }
  return s || undefined;
}
