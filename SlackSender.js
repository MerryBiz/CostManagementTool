// Compiled using ts2gas 3.4.4 (TypeScript 3.7.4)
var exports = exports || {};
var module = module || { exports: exports };
"use strict";
//定数
var SLACK_CHANNEL = "#6cp_admin_it_office";
var BOT_NAME = "原価自動集計Bot";
var MENTIONS = ["@hirokazu.nezu"];
/*
* @param <SalesMeetingMemo> salesMeetingMemo
*/
function sendSlack(errorText) {
    send(generateMessage(errorText));
}
/*
* @param {string} mesage
*/
function send(message) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var url = scriptProperties.getProperty('SLACK_WEBHOOK_URL'); // URLをスクリプトプロパティから取得
    
    var data = { "channel": SLACK_CHANNEL, "username": BOT_NAME, "text": message };
    var payload = JSON.stringify(data);
    
    var options = {
        "method": "POST",
        "contentType": "application/json",
        "payload": payload
    };
    
    var response = UrlFetchApp.fetch(url, options);
}
function generateMessage(text) {
    var message = "";
    for (var i = 0; i < MENTIONS.length; i++) {
        message += "<" + MENTIONS[i] + "> ";
    }
    message += "\n";
    message += "【エラー発生報告】";
    message += "\n";
    message += "\n";
    message += "原価管理自動集計ツールでエラーが発生しました。ログを確認してください。";
    message += "\n";
    message += "---------------------------------------";
    message += "\n";
    message += text;
    message += "\n";
    message += "---------------------------------------";
    message += "\n";
    message += "https://script.google.com/home/projects/1LdAQVjjcqJ4Bc__xie7HxLjL8g2ZTvH0JXgS8Gn6Ht_vO4TH7Exo5F7Z/executions";
    return message;
}
