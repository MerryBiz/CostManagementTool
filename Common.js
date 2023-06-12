// Compiled using ts2gas 3.4.4 (TypeScript 3.7.4)
var exports = exports || {};
var module = module || { exports: exports };
"use strict";

var scriptProperties = PropertiesService.getScriptProperties();


var COST_MANAGEMENT_SHEET_ID_FOR_TEST = scriptProperties.getProperty('COST_MANAGEMENT_SHEET_ID_FOR_TEST');  //原価管理マスタ テスト用
var SALES_MANAGEMENT_SHEET_ID_FOR_TEST = scriptProperties.getProperty('SALES_MANAGEMENT_SHEET_ID_FOR_TEST'); //(新)契約管理シート テスト用
var STAFF_ATTENDANCE_FOLDER_ID_FOR_TEST = scriptProperties.getProperty('STAFF_ATTENDANCE_FOLDER_ID_FOR_TEST');
var CONSULTANT_ATTENDANCE_FOLDER_ID_FOR_TEST = scriptProperties.getProperty('CONSULTANT_ATTENDANCE_FOLDER_ID_FOR_TEST');
var COST_MANAGEMENT_SHEET_ID = scriptProperties.getProperty('COST_MANAGEMENT_SHEET_ID');  //原価管理マスタ
var STAFF_ATTENDANCE_FOLDER_ID = scriptProperties.getProperty('STAFF_ATTENDANCE_FOLDER_ID');//勤務実績表フォルダ
var INACTIVE_STAFF_ATTENDANCE_FOLDER_ID = scriptProperties.getProperty('INACTIVE_STAFF_ATTENDANCE_FOLDER_ID'); //非稼働勤務実績表フォルダ
var SALES_MANAGEMENT_SHEET_ID = scriptProperties.getProperty('SALES_MANAGEMENT_SHEET_ID'); //(新)契約管理シート

var SALES_STOCK_SHEET_NAME = "請求データ連携";
var STAFF_ATTENDANCE_STOCK_SHEET_NAME = "スタッフ勤務実績連携";
var CONSULTANT_ATTENDANCE_STOCK_SHEET_NAME = "コンサルタント勤務実績連携";
var CNT_SUFFIX = "_cnt";
var isTest = false;
function getPrevMonthTitle() {
    var now = new Date();
    now.setMonth(now.getMonth() - 1);
    var prevMonthTitle = now.getFullYear() + "年" + (now.getMonth() + 1) + "月";
    return prevMonthTitle;
}
//原価管理シートのId
function getCostManagementSheetId() {
    if (isTest) {
        return COST_MANAGEMENT_SHEET_ID_FOR_TEST;
    }
    else {
        return COST_MANAGEMENT_SHEET_ID;
    }
}
//契約管理シート
function getSalesManagementSheetId() {
    if (isTest) {
        return SALES_MANAGEMENT_SHEET_ID_FOR_TEST;
        // return SALES_MANAGEMENT_SHEET_ID;
    }
    else {
        return SALES_MANAGEMENT_SHEET_ID;
    }
}
//スタッフ勤務実績フォルダ
function getStaffAttendanceFolderId() {
    if (isTest) {
        return STAFF_ATTENDANCE_FOLDER_ID_FOR_TEST;
        // return STAFF_ATTENDANCE_FOLDER_ID;
    }
    else {
        return STAFF_ATTENDANCE_FOLDER_ID;
    }
}
//非稼働スタッフ勤務実績フォルダ
function getInactiveStaffAttendanceFolderId() {
    if (isTest) {
        return STAFF_ATTENDANCE_FOLDER_ID_FOR_TEST;
        // return STAFF_ATTENDANCE_FOLDER_ID;
    }
    else {
        return INACTIVE_STAFF_ATTENDANCE_FOLDER_ID;
    }
}
//コンサル勤務実績フォルダ
function getConsultantAttendanceFolderId() {
    if (isTest) {
        return CONSULTANT_ATTENDANCE_FOLDER_ID_FOR_TEST;
    }
    else {
        return "";
    }
}
/*
* シート内の特定の列内の文字列を検索する。便利。
* @param <Sheet> sheet 検索対象のシート
* @param <String> val 検索文字列
* @param <int> col 検索列数(ex, A列 = 1)
*
* @return {int} 行数
*/
function findRow(dat, val, col) {
    for (var i = 0; i < dat.length; i++) {
        if (dat[i][col] === val) {
            return i;
        }
    }
    return 0;
}
/*
* シート内の特定の列内の文字列を検索する。便利。
* @param <Sheet> sheet 検索対象のシート
* @param <String> val 検索文字列
* @param <int> row 検索列数(ex, A列 = 1)
*
* @return {int} 行数
*/
function findCol(dat, val, row) {
    for (var i = 0; i < dat.length; i++) {
        if (dat[row][i] === val) {
            return i;
        }
    }
    return 0;
}
function checkCustomerId(customerId) {
    var regex = new RegExp(/^A[0-9]{5}$/);
    if (typeof (customerId) !== "string" || !regex.test(customerId)) {
        console.warn("顧客IDが不正です。顧客ID：" + customerId);
        return false;
    }
    return true;
}
function delete_specific_triggers(name_function) {
    var all_triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < all_triggers.length; ++i) {
        if (all_triggers[i].getHandlerFunction() == name_function)
            ScriptApp.deleteTrigger(all_triggers[i]);
    }
}
function needRestart(start_time, currentCnt, funcName) {
    var current_time = new Date();
    var difference = (current_time.getTime() - start_time.getTime()) / (1000 * 60);
    // = parseInt((current_time.getTime() - start_time.getTime()) / (1000 * 60));
    //4分を超えていたら中断処理
    if (difference >= 4) {
        currentCnt++;
        var properties = PropertiesService.getScriptProperties();
        properties.setProperty(funcName + CNT_SUFFIX, currentCnt.toString());
        ScriptApp
            .newTrigger(funcName)
            .timeBased()
            .everyMinutes(1)
            .create();
        console.log("6 minutes restart!! Next Start is :" + currentCnt);
        return true;
    }
    return false;
}
function removeCollectedRows(targetSheet, targetMonthName) {
    var displayData = targetSheet.getDataRange().getDisplayValues();
    var startRowIdx = 0;
    var endRowIdx = 0;
    var foundThisMonth = false;
    Logger.log("displayDatas.length:" + displayData.length);
    for (var row = 0; row < displayData.length; row++) {
        if (displayData[row][0] === targetMonthName && !foundThisMonth) {
            Logger.log("startRowIdx is setted" + startRowIdx);
            startRowIdx = row;
            foundThisMonth = true;
        }
        if (foundThisMonth && displayData[row][0] !== targetMonthName) {
            Logger.log("endRowIdx is setted" + endRowIdx);
            endRowIdx = row - 1;
            break;
        }
    }
    Logger.log("start:" + startRowIdx + "end" + endRowIdx);
    if (foundThisMonth) {
        if (endRowIdx === 0) {
            endRowIdx = displayData.length - 1;
        }
        console.log((startRowIdx + 1) + "行目から" + (endRowIdx - startRowIdx + 1) + "行削除しました。");
        targetSheet.deleteRows(startRowIdx + 1, endRowIdx - startRowIdx + 1);
    }
}
