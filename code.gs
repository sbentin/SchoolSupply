/**
 * Google Apps Script for the SchoolSupply order sheet.
 * Paste into: spreadsheet > Extensions > Apps Script, replacing doPost.
 * IMPORTANT: after saving, go to Deploy > Manage deployments > ✎ Edit >
 * Version: "New version" > Deploy — otherwise the web app keeps running
 * the old code.
 *
 * To test from the editor, run testOrder() (▶), NOT doPost — doPost only
 * receives data via a real HTTP POST; running it directly has no event
 * object and throws on e.postData.
 *
 * Expected column order: timestamp | name | phone | email | order description | total
 * (adjust the appendRow below if your sheet's columns differ)
 */

// Hebrew labels for the per-item choices sent by the site (cart[].custom).
// Keys not listed here still appear, using the raw key as the label.
var CUSTOM_LABELS = {
  stickers: "מדבקות שם",
  ruler: "סרגל",
  scissors: "מספריים",
  headset: "אוזניות ועכבר אלחוטי",
};

function describeCart(cart) {
  return (cart || []).map(function (item) {
    var custom = item.custom || {};
    var choices = Object.keys(custom).map(function (key) {
      return (CUSTOM_LABELS[key] || key) + ": " + custom[key];
    }).join(" · ");
    var line = "כיתה " + item.letter + " (" + item.price + " ₪)";
    return choices ? line + " — " + choices : line;
  }).join("\n");
}

function doPost(e) {
  if (!e || !e.postData) {
    throw new Error("No POST data — to test from the editor, run testOrder() instead of doPost.");
  }
  var data = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  // 1. הוספת השורה עם הערכים false
  sheet.appendRow([
    new Date(data.timestamp), 
    data.name,
    "'" + data.phone, 
    data.email,
    describeCart(data.cart),
    data.total,
    false, // עמודה G: שולם
    false  // עמודה H: מוכן
  ]);
  
  var lastRow = sheet.getLastRow();
  var checkboxRange = sheet.getRange(lastRow, 7, 1, 2); // שורה אחרונה, עמודה 7 (G), שורה אחת, 2 עמודות (G ו-H)
  var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange.setDataValidation(checkboxValidation);
  
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}