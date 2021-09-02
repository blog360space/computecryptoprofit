/**
 * 1/ Save this document to your google drive: File > Make a coppy.
 * https://docs.google.com/spreadsheets/d/1WMKnPcUdnaUlr9suTfyhgbg1r8wLTEw4OaCBkgyVxtw/edit#gid=0
 * 2/ Go to Tool > Script Sditor
 *    Run function onOPen, then you will have new menu "My menu"
   3/ Click on My menu > Update crypto price (you will need to grant permission to it)
 * 4/ Duplicate example and fill your own data.
 *
 * This spreadsheets was added some formula which use full for me to calculate the profit.
 * Script CryptoReport is to download latest crypto currency price from binance, and the formula will calculate profit.
 */

var CryptoReport = {
    /**
     * Api binance to get price and set to sheet CoinPrice
     */
    apiPriceEndpoint: 'https://api2.binance.com/api/v3/ticker/price',

    /**
     * Base currency USDT is more common than eth or btc
     */
    currency: 'USDT',
    /**
     * max row coin list in sheet CoinPrice
     */
    MAX: 100,

    sheet: null,

    listCryptos: [],

    setData: function() {
        this.initListCrypto();
        this.getPriceFromApi();
        this.updatePriceToSpreadsheet();
    },

    initListCrypto: function() {
        var sheet = this.getSheet();
        var cell = cellText = null;
        for (var i = 1; i <= this.MAX; i++) {
            cell = sheet.getRange("A" + i).getRichTextValue();
            cellText = cell.getText() + this.currency;
            if (cellText === this.currency) {
                return;
            }
            this.listCryptos.push({
                'symbol': cellText.toUpperCase(),
                'price': 0
            });
        }
    },

    /**
     * Get price from api binance
     */
    getPriceFromApi: function() {
        var url = this.apiPriceEndpoint;
        var response = UrlFetchApp.fetch(url);
        var json = response.getContentText();
        var data = JSON.parse(json);
        var count = 0;
        var totalRow = this.listCryptos.length;

        for (var i = 0; i < totalRow; i++) {
            var thisRow = this.listCryptos[i];
            for (var j = 0; j < data.length; j++) {
                var thatRow = data[j];
                if (thisRow.symbol === thatRow.symbol) {
                    this.listCryptos[i].price = thatRow.price;
                    count ++;
                    if (count === totalRow) {
                        return;
                    }
                }
            }
        }
    },

    /**
     * Update Price to sheet CoinPrice
     */
    updatePriceToSpreadsheet: function() {
        var sheet = this.getSheet();
        var totalRow = this.listCryptos.length;
        for (var i = 0; i < totalRow; i ++) {
            var cell = sheet.getRange("B" + (i + 1));
            var thisRow = this.listCryptos[i];
            if (thisRow.price === 0) {
                continue;
            }
            var value = SpreadsheetApp.newRichTextValue()
                .setText(thisRow.price)
                .build();
            cell.setRichTextValue(value);
        }
        SpreadsheetApp.flush();
    },

    /**
     * Get the CoinPrice sheet
     * @returns SpreadsheetApp CoinPrice
     */
    getSheet: function() {
        if (!this.sheet) {
            this.sheet = SpreadsheetApp.getActive().getSheetByName('CoinPrice');
        }
        return this.sheet;
    }
};

function main() {
    CryptoReport.setData();
}

/**
 * Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
 * @param e
 */
function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('My Menu')
        .addItem('Update crypto price', 'main')
        .addToUi();
}
