/**
 * @fileoverview モジュール学習管理 - 定数定義
 * @description モジュール学習管理機能で使用する定数群を定義します。
 */

const MODULE_DEFAULT_CYCLES = [
  { order: 1, startMonth: 6, endMonth: 7, label: '6-7' },
  { order: 2, startMonth: 9, endMonth: 10, label: '9-10' },
  { order: 3, startMonth: 11, endMonth: 12, label: '11-12' },
  { order: 4, startMonth: 1, endMonth: 2, label: '1-2' }
];

const MODULE_DEFAULT_KOMA_PER_CYCLE = 7;
const MODULE_DISPLAY_HEADER = 'MOD実施累計(表示)';
const MODULE_WEEKLY_LABEL = '今週';
const MODULE_GRADE_MIN = 1;
const MODULE_GRADE_MAX = 6;
const MODULE_SETTINGS_PREFIX = 'MODULE_';

const MODULE_WEEKDAY_PRIORITY = {
  1: 0, // 月
  3: 1, // 水
  5: 2, // 金
  2: 3, // 火
  4: 4  // 木
};

const MODULE_CONTROL_MARKERS = {
  PLAN: 'PLAN_TABLE',
  EXCEPTIONS: 'EXCEPTIONS_TABLE'
};

const MODULE_CONTROL_DEFAULT_LAYOUT = {
  VERSION_ROW: 1,
  PLAN_MARKER_ROW: 3,
  EXCEPTIONS_MARKER_ROW: 40
};

const MODULE_CONTROL_PLAN_HEADERS = [
  'fiscal_year',
  'cycle_order',
  'start_month',
  'end_month',
  'g1_koma',
  'g2_koma',
  'g3_koma',
  'g4_koma',
  'g5_koma',
  'g6_koma',
  'note'
];

const MODULE_CONTROL_EXCEPTION_HEADERS = [
  'date',
  'grade',
  'delta_sessions',
  'reason',
  'note'
];
