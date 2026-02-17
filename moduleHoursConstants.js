/**
 * @fileoverview モジュール学習管理 - 定数定義
 * @description モジュール学習管理機能で使用する定数群を定義します。
 */

/** 年間デフォルト目標コマ数（旧: 4クール × 7コマ = 28） */
const MODULE_DEFAULT_ANNUAL_KOMA = 28;

const MODULE_DISPLAY_HEADER = 'MOD実施累計(表示)';
const MODULE_WEEKLY_LABEL = '今週';
const MODULE_RESERVE_LABEL = '予備';
const MODULE_DEFICIT_LABEL = '不足';
const MODULE_GRADE_MIN = 1;
const MODULE_GRADE_MAX = 6;
const MODULE_SETTINGS_PREFIX = 'MODULE_';

const MODULE_WEEKDAY_PRIORITY = Object.freeze({
  1: 0, // 月
  3: 1, // 水
  5: 2, // 金
  2: 3, // 火
  4: 4  // 木
});

/** 実施曜日のデフォルト（月水金 = getDay() の 1,3,5） */
const MODULE_DEFAULT_WEEKDAYS_ENABLED = Object.freeze([1, 3, 5]);

/** 曜日ラベル（getDay()値 → 日本語） */
const MODULE_WEEKDAY_LABELS = Object.freeze({
  1: '月',
  2: '火',
  3: '水',
  4: '木',
  5: '金'
});

const MODULE_CONTROL_MARKERS = Object.freeze({
  PLAN: 'PLAN_TABLE',
  EXCEPTIONS: 'EXCEPTIONS_TABLE'
});

const MODULE_CONTROL_DEFAULT_LAYOUT = Object.freeze({
  VERSION_ROW: 1,
  PLAN_MARKER_ROW: 3,
  EXCEPTIONS_MARKER_ROW: 40
});

/** 計画モード定数 */
const MODULE_PLAN_MODE_ANNUAL = 'annual';
const MODULE_PLAN_MODE_MONTHLY = 'monthly';

/** 年間目標テーブルのヘッダー（V4: 学年行・月別目標対応） */
const MODULE_CONTROL_PLAN_HEADERS = Object.freeze([
  'fiscal_year',
  'grade',
  'plan_mode',
  'm4', 'm5', 'm6', 'm7', 'm8', 'm9',
  'm10', 'm11', 'm12', 'm1', 'm2', 'm3',
  'annual_koma',
  'note'
]);

const MODULE_CONTROL_EXCEPTION_HEADERS = Object.freeze([
  'date',
  'grade',
  'delta_sessions',
  'reason',
  'note'
]);

