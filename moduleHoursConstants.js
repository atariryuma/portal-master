/**
 * @fileoverview モジュール学習管理 - 定数定義
 * @description モジュール学習管理機能で使用する定数群を定義します。
 */

/** 年間デフォルト目標コマ数（旧: 4クール × 7コマ = 28） */
const MODULE_DEFAULT_ANNUAL_KOMA = 28;

const MODULE_DISPLAY_HEADER = 'MOD実施累計(表示)';
const MODULE_WEEKLY_LABEL = '今週';
const MODULE_RESERVE_LABEL = '予備';
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

const MODULE_CONTROL_MARKERS = Object.freeze({
  PLAN: 'PLAN_TABLE',
  EXCEPTIONS: 'EXCEPTIONS_TABLE'
});

const MODULE_CONTROL_DEFAULT_LAYOUT = Object.freeze({
  VERSION_ROW: 1,
  PLAN_MARKER_ROW: 3,
  EXCEPTIONS_MARKER_ROW: 40
});

/** 年間目標テーブルのヘッダー（V3: クール制廃止、年間制へ移行） */
const MODULE_CONTROL_PLAN_HEADERS = Object.freeze([
  'fiscal_year',
  'g1_annual_koma',
  'g2_annual_koma',
  'g3_annual_koma',
  'g4_annual_koma',
  'g5_annual_koma',
  'g6_annual_koma',
  'note'
]);

/** V2→V3 マイグレーション用: 旧クール計画の列数 */
const MODULE_LEGACY_CYCLE_PLAN_COLUMN_COUNT = 11;

const MODULE_CONTROL_EXCEPTION_HEADERS = Object.freeze([
  'date',
  'grade',
  'delta_sessions',
  'reason',
  'note'
]);

