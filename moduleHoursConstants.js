/**
 * @fileoverview モジュール学習管理 - 定数定義
 * @description モジュール学習管理機能で使用する定数群を定義します。
 */

const MODULE_DEFAULT_CYCLES = Object.freeze([
  Object.freeze({ order: 1, startMonth: 6, endMonth: 7, label: '6-7' }),
  Object.freeze({ order: 2, startMonth: 9, endMonth: 10, label: '9-10' }),
  Object.freeze({ order: 3, startMonth: 11, endMonth: 12, label: '11-12' }),
  Object.freeze({ order: 4, startMonth: 1, endMonth: 2, label: '1-2' })
]);

/**
 * 1クール（約2ヶ月）あたりのデフォルトコマ数。
 * 1コマ = 45分 = 15分セッション × 3。
 * 7コマ × 4クール = 年間28コマ（約21時間）が標準的なモジュール学習量。
 */
const MODULE_DEFAULT_KOMA_PER_CYCLE = 7;
const MODULE_DISPLAY_HEADER = 'MOD実施累計(表示)';
const MODULE_WEEKLY_LABEL = '今週';
const MODULE_GRADE_MIN = 1;
const MODULE_GRADE_MAX = 6;
const MODULE_SETTINGS_PREFIX = 'MODULE_';

/**
 * モジュール学習の曜日配分優先度（デフォルト: 月水金のみ）
 * 月→水→金の順で優先配分する。火木にはデフォルトで配分しない。
 * 理由: モジュール学習（15分単位の短時間学習）は隔日配分が効果的。
 * 月水金の3日で十分な配分枠を確保でき、火木は通常授業に専念できる。
 * ダイアログの曜日設定で火木を有効にすれば配分対象に追加可能。
 * @const {Object}
 */
const MODULE_WEEKDAY_PRIORITY = Object.freeze({
  1: 0, // 月
  3: 1, // 水
  5: 2  // 金
});

/**
 * デフォルトの曜日配分順（配列形式）
 * MODULE_WEEKDAY_PRIORITY のキー順を配列化したもの。
 * ダイアログ表示や PropertiesService 保存に使用。
 * @const {Array<number>}
 */
const MODULE_DEFAULT_WEEKDAY_ORDER = Object.freeze([1, 3, 5]);

const MODULE_CONTROL_MARKERS = Object.freeze({
  PLAN: 'PLAN_TABLE',
  EXCEPTIONS: 'EXCEPTIONS_TABLE'
});

const MODULE_CONTROL_DEFAULT_LAYOUT = Object.freeze({
  VERSION_ROW: 1,
  PLAN_MARKER_ROW: 3,
  EXCEPTIONS_MARKER_ROW: 40
});

const MODULE_CONTROL_PLAN_HEADERS = Object.freeze([
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
]);

const MODULE_CONTROL_EXCEPTION_HEADERS = Object.freeze([
  'date',
  'grade',
  'delta_sessions',
  'reason',
  'note'
]);
