// build.js のテスト（生成したpptxの中身を確認する）
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const JSZip = require("jszip");
const { build, main } = require("../build");
const { parse } = require("../lib");

const SHEET = path.join(__dirname, "..", "..", "記入シート.yml");

function sheet() {
  return parse(fs.readFileSync(SHEET, "utf8"));
}

// 生成結果を展開し、スライドごとのXML文字列の配列を返す
async function slides(data) {
  const buf = await build(data).write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
  return Promise.all(names.map((n) => zip.file(n).async("string")));
}
const has = (xml, text) => xml.includes(text);
// 指定した文字が斜体（記入欄）で出力されているか
const italicRun = (xml, text) => new RegExp(`<a:rPr[^>]*\\bi="1"[^>]*>(?:(?!</a:r>).)*<a:t>${text}</a:t>`, "s").test(xml);

test.describe("スライド枚数（追加課題の複数条件網羅）", () => {
  for (const [a, b, n] of [[true, true, 14], [true, false, 14], [false, true, 14], [false, false, 13]]) {
    test.it(`更新=${a}, ユーザ登録=${b} → ${n}枚`, async () => {
      const d = sheet();
      d.追加課題.更新 = a;
      d.追加課題.ユーザ登録 = b;
      const xs = await slides(d);
      assert.equal(xs.length, n);
      const all = xs.join();
      assert.equal(has(all, "従業員情報更新"), a);
      assert.equal(has(all, "ユーザ情報登録"), b);
      // 合計：追加課題あり 11.5分（概要を省くと10分）、なし 10分（省くと8.5分）
      assert.equal(has(xs[1], "約11.5分"), a || b);
      assert.equal(has(xs[1], "約8.5分"), !(a || b));
      assert.equal(has(xs[1], "基本ルート＋追加課題"), a || b);
    });
  }
});

test.describe("記入欄の表示", () => {
  test.it("未記入は斜体、記入済みは通常表示", async () => {
    const d = sheet();
    d.チーム名 = "G班";
    d.メンバー[0].氏名 = "山田";
    d.メンバー[1].氏名 = "";
    const xs = await slides(d);
    assert.ok(has(xs[0], "G班") && !italicRun(xs[0], "G班"));
    assert.ok(has(xs[0], "山田") && !italicRun(xs[0], "山田"));
    assert.ok(italicRun(xs[0], "【記入】氏名2"), "空欄は既定の名前を記入欄で表示");
    assert.ok(has(xs[0], "（リーダー）"));
  });
  test.it("見どころが記入済みなら破線枠を使わない", async () => {
    const d = sheet();
    d.追加課題.見どころ = "【記入】";
    const before = (await slides(d))[7];
    d.追加課題.見どころ = "現在の値が初期表示される";
    const after = (await slides(d))[7];
    assert.ok(has(before, 'prstDash val="dash"'));
    assert.ok(!has(after, 'prstDash val="dash"'));
    assert.ok(!italicRun(after, "現在の値が初期表示される"));
  });
  test.it("追加課題が1つならカードは1枚（幅いっぱい）", async () => {
    const d = sheet();
    d.追加課題.ユーザ登録 = false;
    const xs = await slides(d);
    assert.ok(has(xs[7], "① 従業員情報更新"));
    assert.ok(!has(xs[7], "②"));
  });
});

test.describe("発表の流れの担当欄", () => {
  test.it("振り返りの担当が全員記入済みなら名前を重複なく並べる", async () => {
    const d = sheet();
    Object.assign(d.発表担当, { 学んだこと: "山田", 苦労したこと: "鈴木", 工夫点: "山田" });
    const xs = await slides(d);
    assert.ok(has(xs[1], "山田・鈴木"));
  });
  test.it("1人でも未記入なら記入欄のまま", async () => {
    const d = sheet();
    Object.assign(d.発表担当, { 学んだこと: "山田", 苦労したこと: "鈴木", 工夫点: "【記入】" });
    const xs = await slides(d);
    assert.ok(!has(xs[1], "山田・鈴木"));
    assert.ok(italicRun(xs[1], "【記入】"));
  });
});

test.describe("個人ごとの感想（メンバー数）", () => {
  for (const n of [3, 4]) {
    test.it(`${n}人なら枠も${n}つ`, async () => {
      const d = sheet();
      // 記入シートの人数によらず、先頭の人を複製して n 人にそろえる
      d.メンバー = Array.from({ length: n }, (_, i) => d.メンバー[i] || Object.assign({}, d.メンバー[0], { 氏名: `氏名${i + 1}` }));
      const xs = await slides(d);
      const count = (xs[11].match(/name="感想背景\d"/g) || []).length;
      assert.equal(count, n);
    });
  }
});

test.describe("工夫点の数値", () => {
  test.it("未記入の数値は「—」を記入欄で表示", async () => {
    const d = sheet();
    d.テスト結果 = {};
    const xs = await slides(d);
    assert.ok(italicRun(xs[10], "—"));
  });
});

test.describe("main とコマンド実行", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "empsys-"));
  test.it("main は指定先にpptxを書き出す", async () => {
    const out = path.join(dir, "a.pptx");
    assert.equal(await main(SHEET, out), out);
    assert.ok(fs.statSync(out).size > 10000);
  });
  test.it("コマンド実行：正常なら0で終了", () => {
    const out = path.join(dir, "b.pptx");
    const msg = execFileSync("node", [path.join(__dirname, "..", "build.js"), SHEET, out], { encoding: "utf8" });
    assert.match(msg, /作成しました/);
  });
  test.it("コマンド実行：記入シートに誤りがあれば1で終了し理由を表示", () => {
    const bad = path.join(dir, "bad.yml");
    fs.writeFileSync(bad, fs.readFileSync(SHEET, "utf8").replace("メンバー:\n", "メンバー:\n" + "  - 氏名: 追加\n".repeat(3)));
    assert.throws(() => execFileSync("node", [path.join(__dirname, "..", "build.js"), bad, path.join(dir, "c.pptx")],
      { encoding: "utf8", stdio: "pipe" }), (e) => e.status === 1 && /3〜4件/.test(e.stderr));
  });
});
