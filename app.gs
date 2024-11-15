const MAX_COL_SIZE = 10;

function exportCalendarEventsToSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const calendarList = CalendarApp.getAllCalendars();

  const year = 2024; // 取得したい年を指定
  const month = 11; // 取得したい月（1月なら1）を指定

  let dayRows = [];

  calendarList
    .filter((cal) => {
      return !["誕生日", "日本の祝日", "com"].some((omit) =>
        cal.getName().includes(omit)
      );
    })
    .forEach((calendar) => {
      const sheetName = `${calendar.getName()}【${month}月】`;
      let sheet = spreadsheet.getSheetByName(sheetName);

      // 既存のシートを削除して新規作成
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
      }
      sheet = spreadsheet.insertSheet(sheetName);

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

      let currentRow = 1;

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
            range.setValues([dayRow]);
            range.setFontWeight("bold");
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
              targetCell.setValue(event);
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

      drawBorder(sheet, dayRows);
    });
}

function drawBorder(sheet, dayRows) {
  const lastRow = Math.max(dayRows[dayRows.length - 1] + 5, sheet.getLastRow());
  dayRows.forEach((row) => {
    sheet
      .getRange(row, 1, 1, MAX_COL_SIZE)
      .setBorder(true, false, true, false, null, null); // 週ごとに
  });
  Array.from([3, 5, 7, 9]).forEach((col) => {
    sheet
      .getRange(1, col, lastRow, 1)
      .setBorder(false, true, false, false, null, null); // 曜日ごとに
  });
  sheet
    .getRange(1, 1, lastRow, MAX_COL_SIZE)
    .setBorder(true, true, true, true, null, null); // カレンダー全体
}
