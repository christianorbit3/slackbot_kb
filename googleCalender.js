/**
 * Google Calendar APIを使用してカレンダーの空き枠を取得
 * @param {Array<string>} emailAddresses - メールアドレスの配列
 * @param {number} days - 確認する日数
 * @param {string} startTime - 開始時間（HH:mm形式）
 * @param {string} endTime - 終了時間（HH:mm形式） 
 * @param {Date|string|null} startDate - 開始日（指定がない場合は今日から）
 * @return {Object} 空き枠情報
 */
function analyzeCalendarAvailability(emailAddresses, days, startTime = "09:00", endTime = "18:00", startDate = null) {
  try {
    // 各メールアドレスのカレンダーを取得
    const calendars = emailAddresses.map(email => {
      try {
        return CalendarApp.getCalendarById(email);
      } catch (error) {
        Logger.log(`カレンダー取得エラー (${email}): ${error.message}`);
        return null;
      }
    }).filter(cal => cal !== null);

    if (calendars.length === 0) {
      throw new Error('有効なカレンダーが見つかりませんでした。');
    }

    // 開始日を決定
    let baseDate;
    if (startDate) {
      if (typeof startDate === 'string') {
        baseDate = new Date(startDate);
        if (isNaN(baseDate.getTime())) {
          throw new Error('開始日の形式が正しくありません。');
        }
      } else if (startDate instanceof Date) {
        baseDate = new Date(startDate);
      } else {
        baseDate = new Date(); // 今日
      }
    } else {
      baseDate = new Date(); // 今日
    }

    // 時刻を00:00:00にリセット
    baseDate.setHours(0, 0, 0, 0);
    
    // 時間範囲を設定
    const [startHour, startMinute] = startTime.split(':').map(Number);
    const [endHour, endMinute] = endTime.split(':').map(Number);
    
    // 空き枠情報を格納する配列
    const availability = [];
    
    // 各日について空き枠を確認
    for (let d = 0; d < days; d++) {
      const currentDate = new Date(baseDate.getTime() + d * 24 * 60 * 60 * 1000);
      const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][currentDate.getDay()];
      
      // その日の予定を取得
      const startOfDay = new Date(currentDate);
      startOfDay.setHours(startHour, startMinute, 0, 0);
      
      const endOfDay = new Date(currentDate);
      endOfDay.setHours(endHour, endMinute, 0, 0);
      
      // すべてのカレンダーから予定を取得
      const allEvents = calendars.flatMap(calendar => {
        try {
          const events = calendar.getEvents(startOfDay, endOfDay);
          // 参加が「いいえ」の予定を除外
          return events.filter(event => {
            const myStatus = event.getMyStatus();
            return myStatus !== CalendarApp.GuestStatus.NO;
          });
        } catch (error) {
          Logger.log(`予定取得エラー: ${error.message}`);
          return [];
        }
      });


      // 予定を時間順にソート
      allEvents.sort((a, b) => a.getStartTime() - b.getStartTime());
      
      // 空き枠を計算
      const slots = [];
      let currentTime = new Date(startOfDay);
      let currentSlot = null;
      
      // 30分単位で空き枠をチェック
      while (currentTime < endOfDay) {
        const slotEnd = new Date(currentTime.getTime() + 30 * 60 * 1000);
        
        // この時間枠に予定があるかチェック
        const hasEvent = allEvents.some(event => {
          const eventStart = event.getStartTime();
          const eventEnd = event.getEndTime();
          return (
            (currentTime >= eventStart && currentTime < eventEnd) || // スロット開始時間が予定内
            (slotEnd > eventStart && slotEnd <= eventEnd) || // スロット終了時間が予定内
            (currentTime <= eventStart && slotEnd >= eventEnd) // 予定がスロット内
          );
        });
        
        if (!hasEvent) {
          if (currentSlot === null) {
            // 新しい空き枠の開始
            currentSlot = {
              start: currentTime.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }),
              end: slotEnd.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }),
              duration: '30分'
            };
          } else {
            // 既存の空き枠を延長
            currentSlot.end = slotEnd.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
            currentSlot.duration = calculateDuration(currentSlot.start, currentSlot.end);
          }
        } else if (currentSlot !== null) {
          // 予定がある場合、現在の空き枠を保存して新しい空き枠の準備
          slots.push(currentSlot);
          currentSlot = null;
        }
        
        currentTime = slotEnd;
      }
      
      // 最後の空き枠を保存
      if (currentSlot !== null) {
        slots.push(currentSlot);
      }
      
      // その日の空き枠情報を追加
      availability.push({
        date: currentDate.toLocaleDateString('ja-JP', { month: 'long', day: 'numeric' }),
        dayOfWeek: dayOfWeek,
        fullDate: currentDate.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' }),
        slots: slots
      });
    }
    
    return {
      availability: availability,
      startDate: baseDate.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' }),
      error: null
    };
    
  } catch (error) {
    Logger.log('カレンダー空き枠取得エラー: ' + error.message);
    return {
      availability: [],
      error: error.message
    };
  }
}

/**
 * 時間の差分を計算して表示用の文字列を返す
 * @param {string} start - 開始時間（HH:mm形式）
 * @param {string} end - 終了時間（HH:mm形式）
 * @return {string} 表示用の時間差文字列
 */
function calculateDuration(start, end) {
  const [startHour, startMinute] = start.split(':').map(Number);
  const [endHour, endMinute] = end.split(':').map(Number);
  
  let totalMinutes = (endHour * 60 + endMinute) - (startHour * 60 + startMinute);
  
  if (totalMinutes < 0) {
    totalMinutes += 24 * 60;
  }
  
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  
  if (hours > 0) {
    return `${hours}時間${minutes > 0 ? minutes + '分' : ''}`;
  } else {
    return `${minutes}分`;
  }
}

/**
 * Googleカレンダーにイベントを作成
 * @param {Object} eventJson - イベント情報
 * @return {Object} 作成結果
 */
function createCalendarEvent(eventJson) {
  try {
    // kiko.bandai@buki-ya.comのカレンダーを取得
    const calendar = CalendarApp.getCalendarById('kiko.bandai@buki-ya.com');
    
    if (!calendar) {
      throw new Error('指定されたカレンダーが見つかりませんでした。');
    }

    // 開始日時をパース
    const startDateTime = new Date(eventJson.startDateTime);
    if (isNaN(startDateTime.getTime())) {
      throw new Error('開始日時の形式が正しくありません。');
    }

    // 終了日時を計算（開始時間 + duration分）
    const endDateTime = new Date(startDateTime.getTime() + (eventJson.duration || 30) * 60 * 1000);

    // イベントのオプションを設定
    const eventOptions = {
      description: `Slackボットから自動作成された予定\n作成日時: ${new Date().toLocaleString('ja-JP')}`
    };

    // 招待者がいる場合は追加
    if (eventJson.guestEmail && eventJson.guestEmail.trim() !== '') {
      eventOptions.guests = eventJson.guestEmail;
      eventOptions.sendInvites = true;
    }

    // カレンダーイベントを作成
    const event = calendar.createEvent(
      eventJson.title,
      startDateTime,
      endDateTime,
      eventOptions
    );

    // Eventsシートに記録
    saveEventToSheet(eventJson, event.getId());

    // カレンダーのURLを生成（イベント詳細ページ）
    const calendarUrl = `https://calendar.google.com/calendar/u/0/r/eventedit/${event.getId()}`;

    return {
      success: true,
      eventId: event.getId(),
      eventUrl: calendarUrl,
      message: `イベント「${eventJson.title}」を作成しました。`
    };

  } catch (error) {
    Logger.log('カレンダーイベント作成エラー: ' + error.message);
    throw error;
  }
}

/**
 * イベント情報をEventsシートに保存
 * @param {Object} eventJson - イベント情報
 * @param {string} eventId - カレンダーイベントID
 */
function saveEventToSheet(eventJson, eventId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME_EVENTS);
    
    // Eventsシートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME_EVENTS);
      
      // ヘッダー行を設定
      const headers = [
        'イベントID',
        'タイトル',
        '開始日時',
        '時間（分）',
        '招待者',
        '作成日時',
        'ステータス'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#f0f0f0');
    }

    // 新しい行を追加
    const newRow = [
      eventId,
      eventJson.title,
      eventJson.startDateTime,
      eventJson.duration || 30,
      eventJson.guestEmail || '',
      new Date().toLocaleString('ja-JP'),
      '作成済み'
    ];

    sheet.appendRow(newRow);

  } catch (error) {
    Logger.log('Eventsシートへの保存エラー: ' + error.message);
    // エラーが発生してもカレンダーイベント作成は成功しているので、ログのみ記録
  }
}