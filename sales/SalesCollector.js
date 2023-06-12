// Compiled using ts2gas 3.4.4 (TypeScript 3.7.4)
var exports = exports || {};
var module = module || { exports: exports };
"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var SERVICE_TYPE_NAME_VKA = "1.VKA";
var SERVICE_TYPE_NAME_OT_VKA = "3.OT(VKA)";
var CORRECT_TARGET_AMOUNT_KB = "予定";
// 売上情報を取得するmainメソッド
function collectSalesInfo() {
    try {
        var costManagementSpreadSheet = SpreadsheetApp.openById(getCostManagementSheetId());
        var salesStockSheet = costManagementSpreadSheet.getSheetByName(SALES_STOCK_SHEET_NAME);
        if (!salesStockSheet) {
            console.warn("can't get salesStock sheet");
            return;
        }
        var prevMonthTitle = getPrevMonthTitle();
        removeCollectedRows(salesStockSheet, prevMonthTitle);
        // シート統合により集計廃止
        // var stripeCollector = new StripeSalesCollector();
        // stripeCollector.collect();
        var mfCollector = new MFSalesCollector();
        mfCollector.collect();
    }
    catch (e) {
        sendSlack(e);
    }
}
var AbsSalesCollector = /** @class */ (function () {
    function AbsSalesCollector(sheetName, targetAmountKbIdx, serviceTypeKbIdx, searchStartRowIdx) {
        this.sheetDisplayValues = new Array();
        this.sheetValues = new Array();
        this.sheetName = sheetName;
        this.targetAmountKbIdx = targetAmountKbIdx;
        this.serviceTypeKbIdx = serviceTypeKbIdx;
        this.searchStartRowIdx = searchStartRowIdx;
        var spreadSheet = SpreadsheetApp.openById(getSalesManagementSheetId());
        this.sheet = spreadSheet.getSheetByName(sheetName);
        this.sheetDisplayValues = this.sheet.getDataRange().getDisplayValues();
        this.sheetValues = this.sheet.getDataRange().getValues();
    }
    AbsSalesCollector.prototype.collect = function () {
        var targetMonthColIdx = this.findTargetMonthColIdx();
        Logger.log("targetMonthColIdx:" + targetMonthColIdx);
        if (targetMonthColIdx <= 0) {
            console.warn("契約管理シートに対象月の列が見つかりません。検索対象月：" + getPrevMonthTitle());
        }
        var costManagementSpreadSheet = SpreadsheetApp.openById(getCostManagementSheetId());
        var salesStockSheet = costManagementSpreadSheet.getSheetByName(SALES_STOCK_SHEET_NAME);
        if (!salesStockSheet) {
            Logger.log("can't get salesStock");
            return;
        }
        var prevMonthTitle = getPrevMonthTitle();
        Logger.log("sheetValuesLength" + this.sheetValues.length);
        for (var cnt = this.searchStartRowIdx; cnt < this.sheetValues.length; cnt++) {
            var customerId = this.sheetDisplayValues[cnt][0];
            var customerName = this.sheetDisplayValues[cnt][2];
            Logger.log(customerId);
            if (!customerId && !customerName) {
                console.log("検査終了。検査完了行:" + cnt);
                break;
            }
            Logger.log(customerId);
            if (!checkCustomerId(customerId)) {
                continue;
            }
            var serviceTypeKb = this.sheetDisplayValues[cnt][this.serviceTypeKbIdx];
            Logger.log(serviceTypeKb);
            if (!this.checkServiceTypeIsVKA(serviceTypeKb)) {
                continue;
            }
            var targetAmountKb = this.sheetDisplayValues[cnt][this.targetAmountKbIdx];
            Logger.log(targetAmountKb);
            if (targetAmountKb != CORRECT_TARGET_AMOUNT_KB) {
                continue;
            }
            var amount = this.sheetValues[cnt][targetMonthColIdx];
            if (typeof (amount) !== "number" || amount <= 0) {
                continue;
            }
            var stockData = [prevMonthTitle, customerId, customerName, this.getCollectorTypeStr(), amount];
            salesStockSheet.appendRow(stockData);
        }
    };
    AbsSalesCollector.prototype.getCollectorTypeStr = function () {
        return "";
    };
    AbsSalesCollector.prototype.findTargetMonthColIdx = function () {
        var prevMonthTitle = getPrevMonthTitle();
        Logger.log(prevMonthTitle);
        return findCol(this.sheetDisplayValues, prevMonthTitle, 0);
    };
    AbsSalesCollector.prototype.checkServiceTypeIsVKA = function (serviceTypeKb) {
        return serviceTypeKb == SERVICE_TYPE_NAME_VKA || serviceTypeKb == SERVICE_TYPE_NAME_OT_VKA;
    };
    return AbsSalesCollector;
}());
var StripeSalesCollector = /** @class */ (function (_super) {
    __extends(StripeSalesCollector, _super);
    function StripeSalesCollector() {
        return _super.call(this, "売上管理【Stripe】", 4, 6, 8) || this;
    }
    StripeSalesCollector.prototype.getCollectorTypeStr = function () {
        return "Stripe";
    };
    return StripeSalesCollector;
}(AbsSalesCollector));
var MFSalesCollector = /** @class */ (function (_super) {
    __extends(MFSalesCollector, _super);
    function MFSalesCollector() {
        return _super.call(this, "売上管理", 7, 9, 9) || this;
    }
    MFSalesCollector.prototype.getCollectorTypeStr = function () {
        return "MF or Stripe";
    };
    return MFSalesCollector;
}(AbsSalesCollector));
