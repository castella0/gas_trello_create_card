/**
var TRELLO_APP_KEY = "XXXXXXXXXXXXXXXXXXXXXX";
var TRELLO_TOKEN = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";

var TRELLO_API = {
  create_card : "https://trello.com/1/cards"
};

var TRELLO_LABEL_ID = {
  MAIN: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  NEED_MASTER: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  NEED_BUILD: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  NEED_IMAGE: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  NEED_SWF: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  Task: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  Epic: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
};

var TRELLO_LIST = {
  dev_main: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  dev_sub: "XXXXXXXXXXXXXXXXXXXXXXXXXXXX",
};
**/

var INPUT_START_ROW = 8;
var TITLE_INPUT_COLUMN = "A";
var INPUT_SUPPLEMENTARY_LABEL_COLUMN_START = 2; // "B";
var INPUT_SUPPLEMENTARY_LABEL_COLUMN_NUM = 5; // "F"まで;

function myFunction() {

  var yourSelections = Browser.msgBox("Trelloへのカード追加を実行しますか?", Browser.Buttons.OK_CANCEL);
  if ( "ok" != yourSelections ) {
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // 追加先のリストを入力から決める
  var target_list = null;
  var list_key = sheet.getRange("C1").getValue();
  if (list_key && list_key in TRELLO_LIST) {
    target_list = TRELLO_LIST[list_key];
  }
  if (!target_list) {
    Browser.msgBox("unknown list!!");
    return;
  }

  // 追加するカードのラベルを入力から決める
  var default_label = null;
  var default_label_key = sheet.getRange("C2").getValue();
  if (default_label_key && default_label_key in TRELLO_LABEL_ID) {
    default_label = TRELLO_LABEL_ID[default_label_key];
  }

  var target_row = INPUT_START_ROW;
  while (true) {
    var range = sheet.getRange(TITLE_INPUT_COLUMN + target_row);
    var value = range.getValue();

    if (!value || 50 <= target_row) {
      break;
    }

    var label = default_label;
    // さらに追加でラベルの指定があれば付加する
    range = sheet.getRange(target_row, INPUT_SUPPLEMENTARY_LABEL_COLUMN_START, 1, INPUT_SUPPLEMENTARY_LABEL_COLUMN_NUM);
    var supplementary_label_keys = range.getValues();
    for (var row = 0, row_len = supplementary_label_keys.length; row < row_len; row++) {
      var supplementary_label_key_columns = supplementary_label_keys[row];

      for (var column = 0, column_len = supplementary_label_key_columns.length; column < column_len; column++) {

        if (supplementary_label_key_columns[column] in TRELLO_LABEL_ID) {
          if (label) {
            label += ",";
          }
          label += TRELLO_LABEL_ID[supplementary_label_key_columns[column]];

        }
      }
    }

    var payload = {
      "key": TRELLO_APP_KEY,
      "token": TRELLO_TOKEN,
      "idList": target_list,
      "name": value,
      "idLabels": label
    };
    var option = {
      method: "post",
      payload: payload
    }

    var response = UrlFetchApp.fetch(TRELLO_API.create_card, option);

// Logger.log(response.getContentText("UTF-8"));

    target_row++;
  }
}
