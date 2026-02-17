/**
 * @fileoverview 年間行事予定表への反映機能
 * @description マスターシートのデータを年間行事予定表シートに一括バッチ処理で反映します。
 *              全更新をメモリ上で構築し、setValues() 1回で書き込みます。
 */

function updateAnnualEvents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const masterSheet = ss.getSheetByName(MASTER_SHEET.NAME);
    if (!masterSheet) {
      showAlert('「マスター」シートが見つかりません。', 'エラー');
      return;
    }

    const eventSheet = getAnnualScheduleSheetOrThrow();
    const masterLastRow = masterSheet.getLastRow();
    if (masterLastRow < MASTER_SHEET.DATA_START_ROW) {
      showAlert('マスターシートに反映対象データがありません。', '通知');
      return;
    }

    const masterData = masterSheet.getRange('A' + MASTER_SHEET.DATA_START_ROW + ':' + MASTER_SHEET.DATA_RANGE_END + masterLastRow).getValues();

    const confirmation = ui.alert('確認', '年間行事予定表への更新処理を開始します。続行しますか？', ui.ButtonSet.OK_CANCEL);
    if (confirmation !== ui.Button.OK) {
      return;
    }

    // 年間行事予定表の対象列範囲を一括読み取り
    const eventLastRow = eventSheet.getLastRow();
    const eventInternalCol = ANNUAL_SCHEDULE.INTERNAL_EVENT_COLUMN;
    const eventExternalCol = ANNUAL_SCHEDULE.EXTERNAL_EVENT_COLUMN;
    const eventAttStartCol = ANNUAL_SCHEDULE.ATTENDANCE_START_COLUMN;
    const eventLunchCol = ANNUAL_SCHEDULE.LUNCH_COLUMN;
    const eventDateCol = ANNUAL_SCHEDULE.DATE_INDEX + 1;

    // 校内行事列(D)を一括取得・更新
    const internalValues = eventSheet.getRange(1, eventInternalCol, eventLastRow, 1).getValues();
    // 対外行事列(M)を一括取得・更新
    const externalValues = eventSheet.getRange(1, eventExternalCol, eventLastRow, 1).getValues();
    // 給食列(AA)を一括取得・更新
    const lunchValues = eventSheet.getRange(1, eventLunchCol, eventLastRow, 1).getValues();
    // 校時データ(U:Z, 6列)を一括取得・更新
    const attendanceValues = eventSheet.getRange(1, eventAttStartCol, eventLastRow, ANNUAL_SCHEDULE.ATTENDANCE_COLS).getValues();
    // 日付列(B)を一括取得（1日あたり複数学年行の対応で使用）
    const eventDateValues = eventSheet.getRange(1, eventDateCol, eventLastRow, 1).getValues();
    const dateRowIndicesMap = buildDateRowIndicesMap_(eventDateValues);
    const expectedGradeRowsPerDate = Math.floor(MASTER_SHEET.DATA_COLUMN_COUNT / ANNUAL_SCHEDULE.ATTENDANCE_COLS);

    let updateCount = 0;
    let attendanceRowMismatchCount = 0;

    masterData.forEach(function(row) {
      const dateKey = formatDateKey(row[0]);
      const eventRowIndices = dateRowIndicesMap[dateKey];
      if (!Array.isArray(eventRowIndices) || eventRowIndices.length === 0) {
        return;
      }

      eventRowIndices.forEach(function(rowIndex) {
        internalValues[rowIndex][0] = row[MASTER_SHEET.INTERNAL_EVENT_INDEX];
        externalValues[rowIndex][0] = row[MASTER_SHEET.EXTERNAL_EVENT_INDEX];
        lunchValues[rowIndex][0] = row[MASTER_SHEET.LUNCH_INDEX];
      });

      // 校時データ: マスターの36列(E:AN)から6行分(学年行)を抽出して書き込み
      // マスター1行 = 年間行事予定表の6行(学年行) x 6列(校時列)
      const masterAttendance = row.slice(MASTER_SHEET.DATA_START_COLUMN - 1, MASTER_SHEET.DATA_START_COLUMN - 1 + MASTER_SHEET.DATA_COLUMN_COUNT);
      applyAttendanceForDateRows_(
        attendanceValues,
        eventRowIndices,
        masterAttendance,
        ANNUAL_SCHEDULE.ATTENDANCE_COLS
      );

      if (eventRowIndices.length < expectedGradeRowsPerDate) {
        attendanceRowMismatchCount++;
      }

      updateCount++;
    });

    if (attendanceRowMismatchCount > 0) {
      Logger.log('[WARNING] 年間行事予定表の日付行が不足している日付が' + attendanceRowMismatchCount + '件あります。校時データの一部が未反映の可能性があります。');
    }

    // 一括書き込み
    if (updateCount > 0) {
      eventSheet.getRange(1, eventInternalCol, eventLastRow, 1).setValues(internalValues);
      eventSheet.getRange(1, eventExternalCol, eventLastRow, 1).setValues(externalValues);
      eventSheet.getRange(1, eventLunchCol, eventLastRow, 1).setValues(lunchValues);
      eventSheet.getRange(1, eventAttStartCol, eventLastRow, ANNUAL_SCHEDULE.ATTENDANCE_COLS).setValues(attendanceValues);
    }

    if (typeof hideSheetForNormalUse_ === 'function') {
      hideSheetForNormalUse_(MASTER_SHEET.NAME);
    } else if (!masterSheet.isSheetHidden()) {
      masterSheet.hideSheet();
    }
    if (masterSheet.isSheetHidden()) {
      ui.alert('年間行事のインポート完了に伴い、マスターシートは非表示にしました。今後は「年間行事予定表」シートを直接編集してください。');
    } else {
      ui.alert('年間行事のインポートは完了しました。マスターシートの非表示化はスキップされました（表示中シート制約）。');
    }
  } catch (error) {
    showAlert(error.message || error.toString(), 'エラー');
  }
}

/**
 * 年間行事予定表の日付列から「日付文字列 => 行インデックス配列(0-based)」を構築
 * @param {Array<Array<*>>} dateValues - 日付列データ
 * @return {Object} マップ
 */
function buildDateRowIndicesMap_(dateValues) {
  const map = {};
  for (let i = 0; i < dateValues.length; i++) {
    const key = formatDateKey(dateValues[i][0]);
    if (!key) {
      continue;
    }
    if (!map[key]) {
      map[key] = [];
    }
    map[key].push(i);
  }
  return map;
}

/**
 * 1日分の校時データ（36セル）を対象行へ反映
 * @param {Array<Array<*>>} attendanceValues - 年間行事予定表のU:Zキャッシュ
 * @param {Array<number>} rowIndices - 対象行インデックス（0-based）
 * @param {Array<*>} masterAttendance - マスターの校時36セル
 * @param {number} attendanceCols - 1行あたり校時列数（通常6）
 */
function applyAttendanceForDateRows_(attendanceValues, rowIndices, masterAttendance, attendanceCols) {
  if (!Array.isArray(attendanceValues) ||
      !Array.isArray(rowIndices) ||
      !Array.isArray(masterAttendance)) {
    return;
  }

  const columnCount = Number(attendanceCols);
  if (!Number.isInteger(columnCount) || columnCount <= 0) {
    return;
  }

  const availableRowGroups = Math.floor(masterAttendance.length / columnCount);
  const rowsToApply = Math.min(rowIndices.length, availableRowGroups);
  for (let group = 0; group < rowsToApply; group++) {
    const rowIndex = rowIndices[group];
    if (!attendanceValues[rowIndex]) {
      continue;
    }

    for (let col = 0; col < columnCount; col++) {
      const sourceIndex = (group * columnCount) + col;
      attendanceValues[rowIndex][col] = normalizeAttendanceCellValue_(masterAttendance[sourceIndex]);
    }
  }
}

/**
 * 校時セル値を年間行事予定表向けに正規化
 * @param {*} value - セル値
 * @return {*} 正規化後の値
 */
function normalizeAttendanceCellValue_(value) {
  return /^[月火水木金土日][１-６]$/.test(value) ? '○' : value;
}
