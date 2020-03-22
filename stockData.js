stocks = {
  updateSelectedStocks: function() {
    updateStocks(true);
  },
  updateStocks: function(selectedOnly) {
    // TODO pull Row1 headings, so this can be run in any sheet with tickers in ColA
    var sheet = SpreadsheetApp.getActiveSheet();
    var date = new Date(Date.now()).toLocaleString().split(",")[0];
    var activeRangeList = sheet.getActiveRangeList();
    var tickers;

    if (activeRangeList && selectedOnly) {
      // only updating selected lines
      let ranges = activeRangeList.getRanges();
      ranges.forEach(range => {
        var rowNum = range.getRow();
        var ticker = sheet.getRange(`A${rowNum}`).getValue();
        let info = getStockInfo(ticker);
        let values = Object.values(info);
        values.unshift(date);

        setRowValues(sheet, rowNum, values);
      });
    } else {
      // batch update everything
      setHeaders(sheet);
      var startRow = 2;
      var tickers = sheet.getRange(`A${startRow}:D500`).getValues();

      tickers
        .filter(arr => arr.length)
        .map(item => item[0])
        .forEach((ticker, index) => {
          let rowNum = index + startRow;
          let info = getStockInfo(ticker);
          let values = Object.values(info);
          values.unshift(date);

          setRowValues(sheet, rowNum, values);
        });
    }
  },
  setHeaders: function(sheet, headers) {
    var sampleInfo = getStockInfo("AAPL");
    var headers = Object.keys(sampleInfo);
    headers.unshift("Last Updated");
    setRowValues(sheet, 1, headers);
  },
  setRowValues: function(sheet, rowNum, values) {
    var firstIndex = 6;
    var firstCell = numToSSColumn(firstIndex) + rowNum;
    var lastCell = numToSSColumn(firstIndex + values.length - 1) + rowNum;

    Logger.log(values);
    sheet.getRange(`${firstCell}:${lastCell}`).setValues([values]);
  },
  getStockInfo: function(ticker) {
    var info = {};
    var url = "https://finviz.com/quote.ashx?t=" + ticker;

    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var html = res.getContentText();
    var $ = Cheerio.load(html);

    $(".snapshot-table2 .snapshot-td2-cp").each(function() {
      info[$(this).text()] = $(this)
        .next()
        .text();
    });

    return info;
  },
  numToSSColumn: function(num) {
    var s = "",
      t;

    while (num > 0) {
      t = (num - 1) % 26;
      s = String.fromCharCode(65 + t) + s;
      num = ((num - t) / 26) | 0;
    }
    return s || undefined;
  }
};
