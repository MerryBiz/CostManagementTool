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
function triggerCollectStaffCost() {
    try {
        collectStaffCost();
    }
    catch (e) {
        sendSlack(e);
    }
}
function triggerCollectConsultantCost() {
    try {
        collectConsultantCost();
    }
    catch (e) {
        sendSlack(e);
    }
}
function collectStaffCost() {
    var staffCollector = new StaffCostCollector();
    staffCollector.collect(arguments.callee.name);
}
function collectConsultantCost() {
    var consultantCollector = new ConsultantCostCollector();
    consultantCollector.collect(arguments.callee.name);
}
// Abstract --------------------------------------------------------
var AbsCostCollector = /** @class */ (function () {
    function AbsCostCollector() {
    }
    AbsCostCollector.prototype.collect = function (funcName) {
        Logger.log(funcName);
        console.time("CostCollector");
        var start_time = new Date();
        var properties = PropertiesService.getScriptProperties();
        var currentCntProp = properties.getProperty(funcName + CNT_SUFFIX);
        var currentCnt = 0;
        if (currentCntProp) {
            currentCnt = parseInt(currentCntProp);
        }
        console.log("start count is :" + currentCnt);
        delete_specific_triggers(funcName);
        var costManagementSheet = SpreadsheetApp.openById(getCostManagementSheetId());
        var costSheet = costManagementSheet.getSheetByName(this.getCostManagementStockSheetName());
        if (!costSheet) {
            console.warn("ストックシートが取得できません。実行関数：" + funcName);
            return;
        }
        var prevMonthTitle = getPrevMonthTitle();
        if (currentCnt === 0) {
            removeCollectedRows(costSheet, prevMonthTitle);
        }
        var targetFiles = this.getTargetFiles();
        var inputValues = new Array();
        for (var cnt = currentCnt; cnt < targetFiles.length; cnt++) {
            var file = targetFiles[cnt];
            var currentSpreadsheet = SpreadsheetApp.open(file);
            var currentAttendanceSheet = currentSpreadsheet.getSheetByName(prevMonthTitle);
            if (!currentAttendanceSheet) {
                // notFindSheetCnt++;
                console.log("先月分の勤務シートが見つかりませんでした。処理をスキップします。:" + file.getName());
                continue;
            }
            var extractData = this.extractAttendanceSummary(currentAttendanceSheet);
            //use setValues instead of appendRow
            // for (var i = 0; i < extractData.length; i++) {
            //     var inputData = extractData[i];
            //     var customerId = this.extractCustomerId(inputData);
            //     inputData.unshift(customerId);
            //     inputData.unshift(currentSpreadsheet.getName());
            //     inputData.unshift(prevMonthTitle);
            //     costSheet.appendRow(inputData);
            // }
            for (var i = 0; i < extractData.length; i++) {
                var inputData = extractData[i];
                var customerId = this.extractCustomerId(inputData);
                inputData.unshift(customerId);
                inputData.unshift(currentSpreadsheet.getName());
                inputData.unshift(prevMonthTitle);
                inputValues.push(inputData);
            }
            if (needRestart(start_time, cnt, funcName)) {
                if (inputValues.length > 0) {
                    costSheet.getRange(costSheet.getLastRow() + 1, 1, inputValues.length, inputValues[0].length).setValues(inputValues);
                }
                console.log("Restart!! CurrentCnt is " + cnt);
                return;
            }
        }
        if (inputValues.length > 0) {
            costSheet.getRange(costSheet.getLastRow() + 1, 1, inputValues.length, inputValues[0].length).setValues(inputValues);
        }
        properties.setProperty(funcName + CNT_SUFFIX, "0");
        console.timeEnd("CostCollector");
    };
    AbsCostCollector.prototype.extractAttendanceSummary = function (curentSheet) {
        var dat = curentSheet.getRange(this.getTargetRangePosition()).getValues();
        var extractData = [];
        for (var cnt = 0; cnt < dat.length; cnt++) {
            if (dat[cnt][0]) {
                extractData.push(dat[cnt]);
            }
        }
        return extractData;
    };
    AbsCostCollector.prototype.getTargetFiles = function () {
        console.time("sortTime");
        var targetFolder = DriveApp.getFolderById(this.getTargetFolderId());
        Logger.log(targetFolder.getName());
        var files = targetFolder.searchFiles("title contains '勤務実績表'");
        //各スタッフのスプシ毎の処理
        var filesArray = [];
        //検証用のファイル制限
        // var verificationFileNameList = ["S0003_松尾 綾子様_勤務実績表", "S0006_皆見 佳子様_勤務実績表", "S0018_原田 雅美 様_勤務実績表", "S0021_近藤 昌代様_勤務実績表", "S0004_吉益 美江様_勤務実績表"];
        while (files.hasNext()) {
            var file = files.next();
            //   for (var k = 0; k < verificationFileNameList.length; k++) {
            //     if (file.getName() === verificationFileNameList[k]) {
            filesArray.push(file);
            //   break;
            // }
            //   }
        }
        // 非アクティブスタッフのファイル抽出
        var inactiveFolerId = this.getInactiveTargetFolderId();
        if (inactiveFolerId) {
            console.log("get Inactive file list");
            var inactiveTargetFolder = DriveApp.getFolderById(inactiveFolerId);
            Logger.log(inactiveTargetFolder.getName());
            var files = inactiveTargetFolder.searchFiles("title contains '勤務実績表'");
            while (files.hasNext()) {
                var file = files.next();
                filesArray.push(file);
            }
        }
        filesArray.sort(function (a, b) {
            if (a.getName() > b.getName()) {
                return 1;
            }
            else {
                return -1;
            }
        });
        // TODO Need to remove.
        for (var i = 0; i < filesArray.length; i++) {
            Logger.log(filesArray[i].getName());
        }
        console.timeEnd("sortTime");
        return filesArray;
    };
    AbsCostCollector.prototype.getTargetFolderId = function () {
        return "";
    };
    AbsCostCollector.prototype.getInactiveTargetFolderId = function () {
        return "";
    };
    AbsCostCollector.prototype.getTargetRangePosition = function () {
        return "C7:I21";
    };
    AbsCostCollector.prototype.extractCustomerId = function (currentDataArray) {
        return "";
    };
    AbsCostCollector.prototype.getCostManagementStockSheetName = function () {
        return "";
    };
    return AbsCostCollector;
}());
// For Staff --------------------------------------------------------
var StaffCostCollector = /** @class */ (function (_super) {
    __extends(StaffCostCollector, _super);
    function StaffCostCollector() {
        return _super.call(this) || this;
    }
    StaffCostCollector.prototype.getTargetFolderId = function () {
        return getStaffAttendanceFolderId();
    };
    StaffCostCollector.prototype.getInactiveTargetFolderId = function () {
        return getInactiveStaffAttendanceFolderId();
    };
    StaffCostCollector.prototype.getTargetRangePosition = function () {
        return "C7:I21";
    };
    StaffCostCollector.prototype.extractCustomerId = function (currentDataArray) {
        var customerIdColumn = currentDataArray[0];
        var customerIdCandidate = customerIdColumn.split(" ")[0];
        var regex = new RegExp(/^A[0-9]{5}$/);
        if (typeof (customerIdCandidate) == "string" && regex.test(customerIdCandidate)) {
            return customerIdCandidate;
        }
        return "";
    };
    StaffCostCollector.prototype.getCostManagementStockSheetName = function () {
        return STAFF_ATTENDANCE_STOCK_SHEET_NAME;
    };
    return StaffCostCollector;
}(AbsCostCollector));
// For Consultant --------------------------------------------------------
var ConsultantCostCollector = /** @class */ (function (_super) {
    __extends(ConsultantCostCollector, _super);
    function ConsultantCostCollector() {
        return _super.call(this) || this;
    }
    ConsultantCostCollector.prototype.getTargetFolderId = function () {
        return getConsultantAttendanceFolderId();
    };
    ConsultantCostCollector.prototype.getTargetRangePosition = function () {
        return "C7:I21";
    };
    ConsultantCostCollector.prototype.extractCustomerId = function (currentDataArray) {
        var customerIdColumn = currentDataArray[0];
        var customerIdCandidate = customerIdColumn.split(" ")[0];
        var regex = new RegExp(/^A[0-9]{5}$/);
        if (typeof (customerIdCandidate) == "string" && regex.test(customerIdCandidate)) {
            return customerIdCandidate;
        }
        return "";
    };
    ConsultantCostCollector.prototype.getCostManagementStockSheetName = function () {
        return CONSULTANT_ATTENDANCE_STOCK_SHEET_NAME;
    };
    return ConsultantCostCollector;
}(AbsCostCollector));
