const MAX_COL_SIZE = 10;
const CALENDAR_START_ROW = 3;

function exportCalendarEventsToSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const homeSheet = spreadsheet.getSheetByName("ホーム");

  if (!homeSheet) {
    throw new Error("ホームシートが見つかりません。");
  }

  // A2とB2から年と月を取得
  const year = homeSheet.getRange("A2").getValue();
  const month = homeSheet.getRange("B2").getValue();

  if (!year || !month) {
    throw new Error(
      "ホームシートのA2またはB2に年または月が入力されていません。"
    );
  }

  const calendarList = CalendarApp.getAllCalendars();

  calendarList
    .filter((cal) => {
      return !["誕生日", "日本の祝日", "com"].some((omit) =>
        cal.getName().includes(omit)
      );
    })
    .forEach((calendar) => {
      const sheetName = `${calendar.getName()}【${month}月】`;
      let sheet = spreadsheet.getSheetByName(sheetName);
      const dayRows = [];
      // 既存のシートを削除して新規作成
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
      }
      sheet = spreadsheet.insertSheet(sheetName);

      // タイトル
      sheet
        .getRange(1, 1)
        .setValue(`${month}月 `)
        .setFontSize(21)
        .setFontWeight("bold");
      sheet
        .getRange(1, 2)
        .setValue(`${calendar.getName()} 診療予定`)
        .setFontSize(21)
        .setFontWeight("bold");

      // 曜日名
      [
        { col: 1, dayOfWeek: "月" },
        { col: 3, dayOfWeek: "火" },
        { col: 5, dayOfWeek: "水" },
        { col: 7, dayOfWeek: "木" },
        { col: 9, dayOfWeek: "金" },
      ].forEach((val) => {
        sheet
          .getRange(2, val.col)
          .setValue(val.dayOfWeek)
          .setFontSize(14)
          .setFontWeight("bold")
          .setHorizontalAlignment("right");
      });

      const daysInMonth = new Date(year, month, 0).getDate();
      const startDate = new Date(year, month - 1, 1);
      const endDate = new Date(year, month - 1, daysInMonth);
      const events = calendar.getEvents(startDate, endDate);

      // 日付ごとに予定を整理
      const eventMap = {};
      for (let day = 1; day <= daysInMonth; day++) {
        eventMap[day] = [];
      }
      events.forEach((event) => {
        const eventDate = event.getStartTime();
        const day = eventDate.getDate();
        if (eventMap[day]) {
          eventMap[day].push(event.getTitle());
        }
      });

      let currentRow = CALENDAR_START_ROW;

      // 初週の空白を挿入
      const firstDayOfWeek = new Date(year, month - 1, 1).getDay(); // 0:日曜, ..., 6:土曜
      const initialBlanks = firstDayOfWeek === 0 ? 5 : firstDayOfWeek - 1; // 月曜始まりで調整
      let dayRow = Array(initialBlanks * 2).fill(""); // 初週の空白セルを追加

      let day = 1;

      // カレンダーのデータを出力
      while (day <= daysInMonth) {
        const date = new Date(year, month - 1, day);
        const dayOfWeek = date.getDay();

        // 平日のみ追加（月曜～金曜）
        if (dayOfWeek >= 1 && dayOfWeek <= 5) {
          dayRow.push(day);
          dayRow.push(""); // 各日付セルに2列分のスペースを確保
        }
        day++;
        // 1週間（平日5日分）がそろったら日付行を設定し、次の週に進む
        if (dayOfWeek === 5 || day > daysInMonth) {
          if (dayRow.length > 0) {
            const range = sheet.getRange(currentRow, 1, 1, dayRow.length);
            range.setValues([dayRow]).setFontWeight("bold").setFontSize(11);
            dayRows.push(currentRow);
          }
          currentRow++;

          // 各日付に対応する予定を追加
          const eventsRowStart = currentRow;
          let maxEventRows = 0;

          for (let col = 1; col <= dayRow.length; col += 2) {
            const dayNumber = dayRow[col - 1];
            if (!dayNumber || isNaN(dayNumber)) continue; // 空白セルはスキップ

            const eventsForDay = eventMap[dayNumber] || [];
            eventsForDay.forEach((event, eventIndex) => {
              const targetRow = eventsRowStart + Math.floor(eventIndex / 2); // 2個ごとに次の行
              const targetCol = col + (eventIndex % 2); // 奇数なら次の列に配置
              const targetCell = sheet.getRange(targetRow, targetCol);
              targetCell.setValue(event).setFontSize(11);
              targetCell.setNumberFormat("@STRING@");
            });

            maxEventRows = Math.max(
              maxEventRows,
              Math.ceil(eventsForDay.length / 2)
            );
          }

          // 最低5行確保
          maxEventRows = Math.max(maxEventRows, 5);
          currentRow += maxEventRows;

          dayRow = []; // 次の週のために日付行をリセット
        }
      }

      // 最終的に予定行を最低5行確保
      const finalLastRow = Math.max(
        dayRows[dayRows.length - 1] + 5,
        sheet.getLastRow()
      );
      if (sheet.getLastRow() < finalLastRow) {
        sheet.insertRowsAfter(
          sheet.getLastRow(),
          finalLastRow - sheet.getLastRow()
        );
      }

      const lastRow = Math.max(
        dayRows[dayRows.length - 1] + 5,
        sheet.getLastRow()
      );
      // 罫線を引く
      drawBorder(sheet, dayRows, lastRow);
      // 日付に色付け
      dayRows.forEach((row) => {
        sheet.getRange(row, 1, 1, MAX_COL_SIZE).setBackground("#d9ead3");
      });
      // 検査予定のセル作成
      sheet
        .getRange(lastRow + 1, 1, 1, MAX_COL_SIZE)
        .merge()
        .setValue("検査予定")
        .setBorder(true, true, true, true, null, null)
        .setBackground("#d4cccb")
        .setFontSize(11)
        .setHorizontalAlignment("center");
      // 検査予定入力欄の罫線引く
      sheet
        .getRange(lastRow + 1, 1, 7, MAX_COL_SIZE)
        .setBorder(true, true, true, true, null, null);
    });
}

function drawBorder(sheet, dayRows, lastRow) {
  dayRows.forEach((row) => {
    sheet
      .getRange(row, 1, 1, MAX_COL_SIZE)
      .setBorder(true, false, true, false, null, null); // 週ごとに
  });
  Array.from([3, 5, 7, 9]).forEach((col) => {
    sheet
      .getRange(2, col, lastRow - 1, 1)
      .setBorder(false, true, false, false, null, null); // 曜日ごとに
  });
  sheet
    .getRange(2, 1, lastRow - 1, MAX_COL_SIZE)
    .setBorder(true, true, true, true, null, null); // カレンダー全体
}
