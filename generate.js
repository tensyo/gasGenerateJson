/**
 * 全シートをjsonにしてマイドライブに出力する
 * [シートごとの定義]にシート独自の定義を追記
 */
function generateJson() {
  // フォルダ作成 & json出力
  createJson(createOutputFolder());

  // json出力先フォルダを作成
  function createOutputFolder() {
    var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm");
    // フォルダ名
    var dir_name = "samplesDir";
    var searched = DriveApp.searchFolders('title = "' + dir_name + '"');
    var base_dir = searched.hasNext() ? searched.next() : DriveApp.createFolder(dir_name);

    // フォルダを作成
    return base_dir.createFolder(now);
  }

  // jsonを生成
  function createJson(outputFolder) {
    // 全シートを取得
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i in sheets) {
      var objSheet = sheets[i],
        sheetName = objSheet.getName(),
        sheetId = objSheet.getSheetId();

      // use
      var editFlg = 0,
        dirName = null,
        startRow = 2, // 開始行の位置
        startColumn = 2, // 開始列の位置（key用）
        startColumnNumbers = {}, // 開始列の位置のリスト（値用）
        numColumns = 1, // start_columnから取得する列数
        numRows = objSheet.getLastRow(); // 行数

      Logger.log([sheetId, sheetName]);

      // シートごとの定義
      switch (sheetId) {
        // sample1
        // シートが複数ある場合case増やす
        case 0:
          // 生成するフォルダ名
          dirName = "sample1";
          startColumn = 1;
          // 項目ごとにkeyとその列番号を指定する
          // 指定したkeyでファイルが作成される
          startColumnNumbers = { val1: 2, val2: 3 };
          editFlg = 1;
          break;

        default:
          continue;
      }

      // 編集対象ではない場合
      if (!editFlg) continue;

      // 保存先フォルダを作成
      var outputFolderChild = outputFolder.createFolder(dirName);

      // IDリストを取得
      var keys = objSheet.getSheetValues(startRow, startColumn, numRows, numColumns);
      keys = arrayTrim(keys);

      // 列ごとに値を取得してファイルを保存
      for (var columnCode in startColumnNumbers) {
        var startColumnNo = startColumnNumbers[columnCode],
          values = objSheet.getSheetValues(startRow, startColumnNo, numRows, numColumns);

        // 出力用配列を生成
        var jsonObject = {},
          fileKey = columnCode;
        for (var dictIdNo in keys) {
          var dictId = keys[dictIdNo];
          if (!dictId) continue;

          var dictValue = values[dictIdNo];

          // 不正文字列を置換したい場合ここに記述
          if (dictValue !== "") {
            dictValue = dictValue
              .toString()
              .replace(/"/g, ' \\"')
              .replace(/</g, " < ")
              .replace(/\/>/g, "  />")
              .replace(/\n/g, "\\\n")
              .replace(/\\n/g, "\\\n")
              .replace(/\\\n/g, "\n");
          }
          jsonObject[dictId] = dictValue;
        }
        // 保存
        outputFolderChild.createFile(columnCode + ".json", JSON.stringify(jsonObject));
      }
    }
  }

  // 配列値の空白を除く
  function arrayTrim(array) {
    var newArray = [];
    for (var i in array) {
      var value = array[i];
      // getSheetValues独自の形式を文字列に変換する処理
      if (isArray(value)) value = value[0];
      newArray.push(trim(value));
    }
    return newArray;
  }

  function trim(str) {
    return str.replace(/^\s+|\s+$/g, "");
  }

  // isArray pollyfill
  function isArray(arr) {
    return Object.prototype.toString.call(arr) === "[object Array]";
  }

  // isObject pollyfill
  function isObject(item) {
    return typeof item === "object" && item !== null && !isArray(item);
  }
}
