// lib.js のテスト（命令網羅・判定条件網羅・条件網羅・複数条件網羅）
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { PLACEHOLDER, display, normalize, parse, demoMinutes } = require("../lib");

const SHEET = path.join(__dirname, "..", "..", "記入シート.yml");

// 正しい形の最小データを作る
function base() {
  const member = (n) => ({ 氏名: n, 担当作業: "x", 良かった点: "a", 反省点: "b", 次に活かしたいこと: "c" });
  return {
    チーム名: "F班",
    メンバー: [member("A"), member("B"), member("C")],
    開発環境: [{ 項目: "言語", 内容: "Java" }],
    学んだこと: [{ 観点: "a", 内容: "b" }],
    苦労したこと: [{ 見出し: "a", 何が起きたか: "b", どう解決したか: "c" }],
    工夫点: [{ 見出し: "a", 説明: "b" }],
    企業課題Iへの要望: ["a"],
    来年の授業への改善点: ["b"],
  };
}

function errorsOf(raw) {
  try {
    normalize(raw);
  } catch (e) {
    return e.details || [e.message];
  }
  return [];
}

test.describe("display：記入欄の判定", () => {
  // 判定1: value が undefined / null / 空文字（複数条件をすべて個別に確認）
  test.it("undefined は既定の【記入】になる", () => {
    assert.deepEqual(display(undefined), { text: PLACEHOLDER, fill: true });
  });
  test.it("null は既定の【記入】になる", () => {
    assert.deepEqual(display(null), { text: PLACEHOLDER, fill: true });
  });
  test.it("空白だけの文字列は既定の【記入】になる", () => {
    assert.deepEqual(display("   "), { text: PLACEHOLDER, fill: true });
  });
  // 判定2: fallback の有無
  test.it("fallback を指定すると、その文字を記入欄として返す", () => {
    assert.deepEqual(display("", "【記入】氏名1"), { text: "【記入】氏名1", fill: true });
    assert.deepEqual(display(null, ""), { text: "", fill: true });
  });
  // 判定3: 【記入】で始まるかどうか
  test.it("【記入】で始まる値は記入欄のまま（前後の空白は除く）", () => {
    assert.deepEqual(display("  【記入】例：abc "), { text: "【記入】例：abc", fill: true });
  });
  test.it("記入済みの値は通常表示", () => {
    assert.deepEqual(display("山田"), { text: "山田", fill: false });
  });
  test.it("途中に【記入】を含むだけなら記入済みとみなす", () => {
    assert.deepEqual(display("山田【記入】"), { text: "山田【記入】", fill: false });
  });
  test.it("数値も文字列に変換して扱う", () => {
    assert.deepEqual(display(0), { text: "0", fill: false });
  });
});

test.describe("normalize：全体の形", () => {
  test.it("null はエラー", () => assert.throws(() => normalize(null), /書き方が正しくありません/));
  test.it("配列はエラー", () => assert.throws(() => normalize([]), /書き方が正しくありません/));
  test.it("文字列はエラー", () => assert.throws(() => normalize("abc"), /書き方が正しくありません/));
  test.it("正しい形なら整えたデータを返す", () => {
    const d = normalize(base());
    assert.equal(d.チーム名, "F班");
    assert.equal(d.メンバー.length, 3);
    assert.deepEqual(d.追加課題, { 更新: true, ユーザ登録: true, 見どころ: undefined });
    assert.deepEqual(d.発表担当, {});
    assert.deepEqual(d.テスト結果, {});
  });
});

test.describe("normalize：一覧の件数（checkList）", () => {
  test.it("一覧でない値はエラー", () => {
    const raw = base();
    raw.メンバー = "A";
    assert.match(errorsOf(raw).join(), /「メンバー」は「- 」で始まる一覧/);
  });
  test.it("項目がない場合もエラー", () => {
    const raw = base();
    delete raw.工夫点;
    assert.match(errorsOf(raw).join(), /「工夫点」は「- 」で始まる一覧/);
  });
  test.it("下限未満（メンバー2名）はエラー", () => {
    const raw = base();
    raw.メンバー.pop();
    assert.match(errorsOf(raw).join(), /「メンバー」は3〜4件.*現在2件/);
  });
  test.it("上限超え（メンバー5名）はエラー", () => {
    const raw = base();
    raw.メンバー.push(raw.メンバー[0], raw.メンバー[0]);
    assert.match(errorsOf(raw).join(), /現在5件/);
  });
  test.it("下限ちょうど・上限ちょうどは正常", () => {
    const raw = base();
    assert.deepEqual(errorsOf(raw), []);
    raw.メンバー.push(raw.メンバー[0]);
    raw.学んだこと = [1, 2, 3].map(() => ({ 観点: "a", 内容: "b" }));
    assert.deepEqual(errorsOf(raw), []);
  });
  test.it("要望の一覧が空ならエラー", () => {
    const raw = base();
    raw.来年の授業への改善点 = [];
    assert.match(errorsOf(raw).join(), /「来年の授業への改善点」は1〜4件/);
  });
  test.it("複数のエラーはまとめて返す", () => {
    const raw = base();
    raw.メンバー = [];
    raw.工夫点 = [];
    assert.equal(errorsOf(raw).length, 2);
  });
});

test.describe("normalize：一覧の要素（checkItems）", () => {
  for (const [label, item] of [["null", null], ["文字列", "abc"], ["配列", ["a"]]]) {
    test.it(`要素が${label}ならエラー`, () => {
      const raw = base();
      raw.学んだこと = [item];
      assert.match(errorsOf(raw).join(), /「学んだこと」の1件目/);
    });
  }
  // 要望（checkTexts）：要素が オブジェクト／配列ならエラー、null・数値・文字列はよい
  for (const [label, item] of [["オブジェクト", { 項目: "a" }], ["配列", ["a"]]]) {
    test.it(`要望の要素が${label}ならエラー`, () => {
      const raw = base();
      raw.来年の授業への改善点 = ["a", item];
      assert.match(errorsOf(raw).join(), /「来年の授業への改善点」の2件目は「- 」の後ろに文章だけ/);
    });
  }
  test.it("要望の要素が空欄（null）や数値でもよい", () => {
    const raw = base();
    raw.企業課題Iへの要望 = [null, 1];
    assert.deepEqual(errorsOf(raw), []);
  });
  test.it("要望の一覧は文字列の要素でよい", () => {
    const raw = base();
    raw.企業課題Iへの要望 = ["a", "b"];
    raw.来年の授業への改善点 = ["c"];
    assert.deepEqual(errorsOf(raw), []);
  });
});

test.describe("normalize：追加課題（toImplemented とオブジェクト判定）", () => {
  const withExtra = (extra) => Object.assign(base(), { 追加課題: extra });
  test.it("追加課題が文字列なら既定値（両方 true）", () => {
    assert.deepEqual(normalize(withExtra("x")).追加課題, { 更新: true, ユーザ登録: true, 見どころ: undefined });
  });
  test.it("中の機能が文字列なら既定値 true", () => {
    const d = normalize(withExtra({ 従業員情報更新: "x", ユーザ情報登録: 1 }));
    assert.equal(d.追加課題.更新, true);
    assert.equal(d.追加課題.ユーザ登録, true);
  });
  test.it("実装した: null は true", () => {
    assert.equal(normalize(withExtra({ 従業員情報更新: { 実装した: null } })).追加課題.更新, true);
  });
  // 複数条件網羅：(更新, ユーザ登録) の組み合わせ4通り
  for (const [a, b] of [[true, true], [true, false], [false, true], [false, false]]) {
    test.it(`実装した: 更新=${a}, ユーザ登録=${b}`, () => {
      const d = normalize(withExtra({ 従業員情報更新: { 実装した: a }, ユーザ情報登録: { 実装した: b }, 見どころ: "x" }));
      assert.deepEqual(d.追加課題, { 更新: a, ユーザ登録: b, 見どころ: "x" });
    });
  }
  test.it("true/false 以外はエラー", () => {
    const errs = errorsOf(withExtra({ 従業員情報更新: { 実装した: "はい" }, ユーザ情報登録: { 実装した: 0 } }));
    assert.equal(errs.length, 2);
    assert.match(errs[0], /従業員情報更新 > 実装した/);
  });
});

test.describe("normalize：テスト結果・発表担当", () => {
  test.it("オブジェクトならそのまま使う", () => {
    const raw = Object.assign(base(), { テスト結果: { OK: 1 }, 発表担当: { 要望: "A" } });
    const d = normalize(raw);
    assert.deepEqual(d.テスト結果, { OK: 1 });
    assert.deepEqual(d.発表担当, { 要望: "A" });
  });
  test.it("オブジェクトでなければ空として扱う", () => {
    const d = normalize(Object.assign(base(), { テスト結果: 5, 発表担当: "A" }));
    assert.deepEqual(d.テスト結果, {});
    assert.deepEqual(d.発表担当, {});
  });
});

test.describe("parse", () => {
  test.it("同梱の記入シートは正しく読める", () => {
    const d = parse(fs.readFileSync(SHEET, "utf8"));
    assert.equal(d.メンバー.length, 4);
    assert.equal(d.苦労したこと.length, 3);
  });
  test.it("YAMLの書き間違いは行番号付きのエラー", () => {
    assert.throws(() => parse("チーム名: F班\nメンバー:\n  - 氏名: A\n 担当: x"), /行目付近の書き方が正しくありません/);
  });
  test.it("空のファイルは位置情報のない読み込みエラーとして扱う", () => {
    assert.throws(() => parse(""), /記入シートのどこかの書き方/);
  });
  test.it("中身が一覧だけのファイルは形のエラー", () => {
    assert.throws(() => parse("- a\n- b"), /見本の形を崩さない/);
  });
});

test.describe("demoMinutes（複数条件網羅）", () => {
  const cases = [[true, true, 4], [true, false, 4], [false, true, 4], [false, false, 2.5]];
  for (const [a, b, want] of cases) {
    test.it(`更新=${a}, ユーザ登録=${b} → ${want}分`, () => {
      assert.equal(demoMinutes({ 更新: a, ユーザ登録: b }), want);
    });
  }
});
