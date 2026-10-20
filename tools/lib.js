// 記入シート（YAML）の読み込み・検証と、記入欄の判定を行う関数群
"use strict";

const yaml = require("js-yaml");

const PLACEHOLDER = "【記入】";

// 各一覧の上限件数（スライドの枠の数）
const LIMITS = {
  メンバー: { min: 3, max: 4 },
  学んだこと: { min: 1, max: 3 },
  苦労したこと: { min: 1, max: 3 },
  工夫点: { min: 1, max: 3 },
  企業課題Iへの要望: { min: 1, max: 4 },
  来年の授業への改善点: { min: 1, max: 4 },
  開発環境: { min: 1, max: 6 },
};

// 表示用の文字列と、記入欄として目立たせるかどうかを返す
function display(value, fallback) {
  if (value === undefined || value === null || String(value).trim() === "") {
    return { text: fallback === undefined ? PLACEHOLDER : fallback, fill: true };
  }
  const text = String(value).trim();
  return { text, fill: text.startsWith(PLACEHOLDER) };
}

// 一覧の件数を検証する。エラーがあれば errors に追加する
function checkList(errors, name, list) {
  const limit = LIMITS[name];
  if (!Array.isArray(list)) {
    errors.push(`「${name}」は「- 」で始まる一覧で書いてください。`);
    return [];
  }
  if (list.length < limit.min || list.length > limit.max) {
    errors.push(`「${name}」は${limit.min}〜${limit.max}件で書いてください（現在${list.length}件）。`);
  }
  return list;
}

// 一覧の要素が項目を持つ形（「項目名: 値」）になっているか検証する
function checkItems(errors, name, list) {
  list.forEach((item, i) => {
    if (item === null || typeof item !== "object" || Array.isArray(item)) {
      errors.push(`「${name}」の${i + 1}件目の書き方が正しくありません。見本と同じ項目名で書いてください。`);
    }
  });
}

// 文字だけの一覧（要望）の要素が「項目名: 値」や入れ子の一覧になっていないか検証する
function checkTexts(errors, name, list) {
  list.forEach((item, i) => {
    if (item !== null && typeof item === "object") {
      errors.push(`「${name}」の${i + 1}件目は「- 」の後ろに文章だけを書いてください。`);
    }
  });
}

// 「実装した」の値を true / false に変換する（未記入は true とみなす）
function toImplemented(errors, name, value) {
  if (value === undefined || value === null) return true;
  if (value === true || value === false) return value;
  errors.push(`「追加課題 > ${name} > 実装した」は true か false で書いてください。`);
  return true;
}

// 読み込んだデータを検証し、生成に使う形に整える
function normalize(raw) {
  const errors = [];
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error("記入シートの書き方が正しくありません。見本の形を崩さないようにしてください。");
  }
  const lists = {};
  for (const name of Object.keys(LIMITS)) {
    lists[name] = checkList(errors, name, raw[name]);
    if (name === "企業課題Iへの要望" || name === "来年の授業への改善点") checkTexts(errors, name, lists[name]);
    else checkItems(errors, name, lists[name]);
  }
  const extra = raw.追加課題 && typeof raw.追加課題 === "object" ? raw.追加課題 : {};
  const upd = extra.従業員情報更新 && typeof extra.従業員情報更新 === "object" ? extra.従業員情報更新 : {};
  const usr = extra.ユーザ情報登録 && typeof extra.ユーザ情報登録 === "object" ? extra.ユーザ情報登録 : {};
  const test = raw.テスト結果 && typeof raw.テスト結果 === "object" ? raw.テスト結果 : {};
  const owner = raw.発表担当 && typeof raw.発表担当 === "object" ? raw.発表担当 : {};

  const data = {
    チーム名: raw.チーム名,
    発表日: raw.発表日,
    メンバー: lists.メンバー,
    発表担当: owner,
    ログインID: raw.デモのログインID,
    開発環境: lists.開発環境,
    追加課題: {
      更新: toImplemented(errors, "従業員情報更新", upd.実装した),
      ユーザ登録: toImplemented(errors, "ユーザ情報登録", usr.実装した),
      見どころ: extra.見どころ,
    },
    学んだこと: lists.学んだこと,
    苦労したこと: lists.苦労したこと,
    テスト結果: test,
    工夫点: lists.工夫点,
    企業課題Iへの要望: lists.企業課題Iへの要望,
    来年の授業への改善点: lists.来年の授業への改善点,
  };
  if (errors.length > 0) {
    const err = new Error("記入シートに直してほしい箇所があります。\n- " + errors.join("\n- "));
    err.details = errors;
    throw err;
  }
  return data;
}

// 「項目名:」の直後の全角スペースを半角スペースに置き換える（よくある入力ミスの救済）
function fixFullWidthSpace(text) {
  return text.replace(/^([ ]*(?:- )?[^\s#:][^:\n]*):\u3000+/gm, "$1: ");
}

// YAML文字列を読み込んで検証する
function parse(text) {
  let raw;
  try {
    raw = yaml.load(fixFullWidthSpace(text));
  } catch (e) {
    const line = e.mark ? `${e.mark.line + 1}行目付近` : "どこか";
    throw new Error(`記入シートの${line}の書き方が正しくありません。「項目名: 値」の「: 」の後ろに半角スペースがあるか、字下げがずれていないか確認してください。`);
  }
  return normalize(raw);
}

// 追加課題の実装状況から、デモの時間（分）を返す
function demoMinutes(extra) {
  return extra.更新 || extra.ユーザ登録 ? 4 : 2.5;
}

module.exports = { PLACEHOLDER, LIMITS, display, fixFullWidthSpace, normalize, parse, demoMinutes };
