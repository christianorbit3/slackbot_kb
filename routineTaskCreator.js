/**
 * 定例タスクを登録する
 * @param {string} sheetId - タスクシートID
 * @return {boolean} 登録結果
 */
function registerRoutineTasks(sheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const routineSheet = spreadsheet.getSheetByName("Routine");
    
    if (!routineSheet) {
      Logger.log(`Routineシートが見つかりません: ${sheetId}`);
      return false;
    }

    const data = routineSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log(`Routineシートにデータがありません: ${sheetId}`);
      return false;
    }

    // ヘッダー行を取得して列インデックスマップを作成
    const headerRow = data[0];
    const columnMap = {
      summary: headerRow.indexOf("概要"),
      detail: headerRow.indexOf("詳細"),
      frequency: headerRow.indexOf("登録頻度"),
      dayOfWeek: headerRow.indexOf("発行曜日"),
      assignee: headerRow.indexOf("アサイン"),
      lastRegistered: headerRow.indexOf("最終登録日"),
      biweeklyPattern: headerRow.indexOf("隔週パターン")
    };

    // 必須列の存在確認
    if (columnMap.summary === -1 || columnMap.frequency === -1 || 
        columnMap.dayOfWeek === -1 || columnMap.assignee === -1) {
      Logger.log(`Routineシートの必須列が見つかりません: ${sheetId}`);
      return false;
    }

    // 最終登録日列が存在しない場合は追加
    if (columnMap.lastRegistered === -1) {
      const lastColumn = routineSheet.getLastColumn();
      routineSheet.getRange(1, lastColumn + 1).setValue("最終登録日");
      columnMap.lastRegistered = lastColumn;
    }

    // 隔週パターン列が存在しない場合は追加
    if (columnMap.biweeklyPattern === -1) {
      const lastColumn = routineSheet.getLastColumn();
      routineSheet.getRange(1, lastColumn + 1).setValue("隔週パターン");
      columnMap.biweeklyPattern = lastColumn;
    }

    const today = new Date();
    const dayOfWeek = today.getDay(); // 0:日曜日, 1:月曜日, ..., 6:土曜日
    const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
    const currentDayName = dayNames[dayOfWeek];
    Logger.log(currentDayName);

    // 現在の週番号を取得（年間通し）
    const currentWeekNumber = getWeekNumber(today);
    const currentPattern = currentWeekNumber % 2 === 1 ? "A" : "B"; // 奇数週がA、偶数週がB
    Logger.log(currentPattern);

    // タスクシートを取得
    const taskSheet = spreadsheet.getSheetByName(SHEET_NAME_TASKS);
    if (!taskSheet) {
      Logger.log(`タスクシートが見つかりません: ${sheetId}`);
      return false;
    }

    // タスクシートのヘッダー行を取得
    const taskHeaderRow = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
    const taskColumnMap = createColumnIndexMap(taskHeaderRow);

    // 2行目から処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const frequency = String(row[columnMap.frequency]).trim();
      const targetDayOfWeek = String(row[columnMap.dayOfWeek]).trim();
      const biweeklyPattern = columnMap.biweeklyPattern !== -1 ? String(row[columnMap.biweeklyPattern]).trim() : "";
      Logger.log(row);
      Logger.log(targetDayOfWeek);
      // 現在の曜日と一致する場合のみ処理
      if (targetDayOfWeek === currentDayName) {
        // 月次タスクの場合、月初（1日〜7日）以外はスキップ
        if (frequency === "月次" && (today.getDate() < 1 || today.getDate() > 7)) {
          Logger.log(`月次タスク「${row[columnMap.summary]}」は月初以外のためスキップします`);
          continue;
        }

        // 隔週タスクの場合、パターンが一致しない場合はスキップ
        if (frequency === "隔週" && biweeklyPattern !== currentPattern) {
          Logger.log(`隔週タスク「${row[columnMap.summary]}」は現在のパターン(${currentPattern})と一致しないためスキップします`);
          continue;
        }
        Logger.log("aaaa");

        // 最終登録日をチェック
        const lastRegistered = row[columnMap.lastRegistered];
        if (lastRegistered) {
          const lastRegisteredDate = new Date(lastRegistered);
          const daysSinceLastRegistration = Math.floor((today - lastRegisteredDate) / (1000 * 60 * 60 * 24));
          
          // 5日以内に登録済みの場合はスキップ
          if (daysSinceLastRegistration < 5) {
            Logger.log(`タスク「${row[columnMap.summary]}」は${daysSinceLastRegistration}日前に登録済みのためスキップします`);
            continue;
          }
        }

        // 期日を設定
        let dueDate;
        if (frequency === "週次") {
          dueDate = new Date(today);
          dueDate.setDate(today.getDate() + 7);
        } else if (frequency === "月次") {
          dueDate = new Date(today.getFullYear(), today.getMonth() + 1, today.getDate());
        } else if (frequency === "隔週") {
          dueDate = new Date(today);
          dueDate.setDate(today.getDate()-dayOfWeek+dayNames.indexOf(targetDayOfWeek));
        } else {
          continue; // 不明な頻度はスキップ
        }

        // タスク情報を作成
        const taskJson = {
          "概要": String(row[columnMap.summary]).trim(),
          "詳細": columnMap.detail !== -1 ? String(row[columnMap.detail]).trim() : "",
          "期日": Utilities.formatDate(dueDate, "JST", "yyyy-MM-dd"),
          "アサイン": String(row[columnMap.assignee]).trim(),
          "ステータス": "",
          SheetId: sheetId
        };

        // タスクを登録
        const success = createTaskToSheet(taskJson);
        
        // タスク登録が成功した場合、最終登録日を更新
        if (success) {
          routineSheet.getRange(i + 1, columnMap.lastRegistered + 1)
            .setValue(Utilities.formatDate(today, "JST", "yyyy-MM-dd"));
        }
      }
    }

    return true;
  } catch (error) {
    logError(error, "定例タスクの登録中にエラーが発生", { sheetId });
    return false;
  }
}

/**
 * 日付から年間通しの週番号を取得
 * @param {Date} date - 日付
 * @return {number} 週番号（1-53）
 */
function getWeekNumber(date) {
  // 年の最初の日を取得
  const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
  
  // 年の最初の日から何日経過したかを計算
  const pastDaysOfYear = (date - firstDayOfYear) / 86400000;
  
  // 週番号を計算（1から始まる）
  return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
}

/**
 * 定例タスク登録のテスト関数
 */
function testRegisterRoutineTasks() {
  const testSheetId = "1BH4ErMZvvgG9cmk03APwoNmU-cbf2nwUjCFFdjjPk74"; // 実際のテスト用スプレッドシートIDに置き換えてください

  registerRoutineTasks(testSheetId);
}