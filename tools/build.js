// 記入シート（YAML）から成果発表会の資料（pptx）を生成する
//   使い方: node build.js [記入シートのパス] [出力するpptxのパス]
"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { parse, display, demoMinutes } = require("./lib");

const ROOT = path.resolve(__dirname, "..");
const DEFAULT_INPUT = path.join(ROOT, "記入シート.yml");
const DEFAULT_OUTPUT = path.join(ROOT, "成果発表会資料_テンプレート.pptx");

// 配色
const COL = {
  text: "1F2933", white: "FFFFFF", navy: "0B3C49", pale: "EEF4F5", teal: "0F7C8C", orange: "C2410C",
  green: "5B8C5A", purple: "7A6FB0", gold: "D9A43B", gray: "64748B", line: "CBD5DA", lightText: "CFE3E6",
  fillOnDark: "FDBA74", fillBg: "FFFAF6", tagBg: "FFF3EC",
};
const FONT = "Meiryo UI";

function build(data) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625 inch
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "成果発表会資料（従業員情報管理システム）";
  pres.subject = "プロジェクト型演習（JavaWeb）成果発表会";
  pres.author = display(data.チーム名, "").text;

  // ---- レイアウト定義 ----
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: COL.navy },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 8.8, h: 1.2,
        fontSize: 36, bold: true, color: COL.white, valign: "middle" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: COL.white },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 6.8, h: 0.6,
        fontSize: 24, bold: true, color: COL.navy, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "プロジェクト型演習（JavaWeb）成果発表会　｜　従業員情報管理システム empsys",
        options: { x: 0.5, y: 5.25, w: 7.5, h: 0.25, fontSize: 9, color: COL.gray, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.22, w: 0.5, h: 0.3, fontSize: 9, color: COL.gray, align: "right" },
  });

  // ---- 共通部品 ----
  // 値を文字列ランに変換する。未記入は橙の斜体で表示する
  function run(value, opts = {}, fallback) {
    const d = display(value, fallback);
    const o = Object.assign({}, opts);
    if (d.fill) Object.assign(o, { color: opts.fillColor || COL.orange, italic: true });
    delete o.fillColor;
    return { text: d.text, options: o };
  }
  function txt(slide, value, opts, fallback) {
    const r = run(value, { color: opts.color, fillColor: opts.fillColor }, fallback);
    const o = Object.assign({ margin: 0, isTextBox: true }, opts, r.options);
    delete o.fillColor;
    slide.addText(r.text, o);
  }
  function sub(slide, text) {
    slide.addText(text, { x: 0.5, y: 0.9, w: 9.0, h: 0.35, fontSize: 13, color: COL.gray,
      margin: 0, isTextBox: true, objectName: "サブタイトル" });
  }
  function tag(slide, label, value) {
    const r = run(value, { color: COL.orange });
    slide.addText([{ text: label, options: { color: COL.orange } }, { text: r.text, options: r.options }], {
      shape: pres.ShapeType.roundRect, rectRadius: 0.08, x: 7.4, y: 0.38, w: 2.1, h: 0.42, fontSize: 11, bold: true,
      fill: { color: COL.tagBg }, line: { color: COL.orange, width: 1 }, align: "center", valign: "middle",
      isTextBox: true, objectName: "担当タグ" });
  }
  function badge(slide, x, y, label, color, size = 0.42) {
    slide.addText(String(label), { shape: pres.ShapeType.ellipse, x, y, w: size, h: size, fill: { color },
      color: COL.white, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
      objectName: "番号バッジ" + label });
  }
  function panel(slide, x, y, w, h, name) {
    slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: COL.pale },
      line: { color: COL.pale }, objectName: name });
  }
  function bulletRuns(values, opts) {
    return values.map((v, i) => {
      const r = run(v, opts);
      return { text: r.text, options: Object.assign({ bullet: true, breakLine: i < values.length - 1 }, opts, r.options) };
    });
  }
  function card(slide, x, y, w, h, head, items, color, name) {
    panel(slide, x, y, w, h, name + "_背景");
    slide.addText(head, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.35, fontSize: 14, bold: true, color,
      margin: 0, isTextBox: true, objectName: name + "_見出し" });
    slide.addText(bulletRuns(items, { color: COL.text }), { x: x + 0.15, y: y + 0.48, w: w - 0.3, h: h - 0.58,
      fontSize: 12, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true, objectName: name + "_本文" });
  }

  const members = data.メンバー;
  const owner = data.発表担当;
  const extra = data.追加課題;
  const hasExtra = extra.更新 || extra.ユーザ登録;
  const memberNames = members.map((m, i) => run(m.氏名, {}, `【記入】氏名${i + 1}`));

  // ===== 1. 表紙 =====
  pres.addSection({ title: "表紙" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "表紙" });
  s.addText("プロジェクト型演習（JavaWeb）成果発表会", { x: 0.6, y: 1.0, w: 8.8, h: 0.4, fontSize: 14,
    color: COL.lightText, margin: 0, isTextBox: true, objectName: "講座名" });
  s.addText("従業員情報管理システム empsys", { placeholder: "title" });
  const coverLine = (label, valueRuns, y, name, h = 0.4) => {
    s.addText([{ text: label, options: { color: COL.white, bold: true } }].concat(valueRuns),
      { x: 0.6, y, w: 8.8, h, fontSize: 16, valign: "top", margin: 0, isTextBox: true, objectName: name });
  };
  const onDark = (r) => ({ text: r.text, options: Object.assign({}, r.options, r.options.italic ? { color: COL.fillOnDark } : { color: COL.white }) });
  const memberRuns = [];
  memberNames.forEach((r, i) => {
    if (i > 0) memberRuns.push({ text: "／", options: { color: COL.white } });
    memberRuns.push(onDark(r));
    if (i === 0) memberRuns.push({ text: "（リーダー）", options: { color: COL.white } });
  });
  coverLine("チーム名：", [onDark(run(data.チーム名))], 3.0, "チーム名");
  coverLine("メンバー：", memberRuns, 3.45, "メンバー", 0.75);
  coverLine("発表日：", [onDark(run(data.発表日))], 4.25, "発表日");
  s.addText("この資料は記入シート.yml から自動で作られています。内容の修正は記入シートで行ってください。", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.3, fontSize: 10, color: "9FC3C9", margin: 0, isTextBox: true, objectName: "注意" });
  s.addNotes("【担当：リーダー／目安 0.5分】\nチーム名とメンバーを紹介する。全員がホワイトボード前に並び、一人ずつ名乗る。");

  // ===== 2. 発表の流れ =====
  pres.addSection({ title: "導入" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "導入" });
  s.addText("発表の流れ", { placeholder: "title" });
  sub(s, "1チーム 8〜12分。全員がホワイトボード前に出て発表します");
  const demoMin = demoMinutes(extra);
  const flow = [
    ["システム概要", 1.5, owner.システム概要, "※最初のチームのみ発表"],
    ["デモンストレーション", demoMin, owner.デモ操作, hasExtra ? "基本ルート＋追加課題" : "基本ルート"],
    ["学んだこと／苦労したこと／工夫点", 3, [owner.学んだこと, owner.苦労したこと, owner.工夫点], ""],
    ["個人ごとの感想", 2, "全員", "1人30秒程度"],
    ["今後の授業への要望", 1, owner.要望, ""],
  ];
  const head = (t, align) => ({ text: t, options: { bold: true, color: COL.white, fill: { color: COL.navy }, align } });
  const rows = [[head("No", "center"), head("内容"), head("目安", "center"), head("担当"), head("備考")]];
  flow.forEach((f, i) => {
    let who;
    if (Array.isArray(f[2])) {
      const names = [...new Set(f[2].map((v) => display(v).text))];
      const allFilled = f[2].every((v) => !display(v).fill);
      who = allFilled ? { text: names.join("・") } : run(undefined);
    } else {
      who = run(f[2]);
    }
    rows.push([{ text: String(i + 1), options: { align: "center" } }, { text: f[0] },
      { text: f[1] + "分", options: { align: "center" } }, who, { text: f[3], options: { color: COL.gray } }]);
  });
  const total = flow.reduce((a, f) => a + f[1], 0);
  rows.push([{ text: "" }, { text: "合計", options: { bold: true } },
    { text: `約${total}分`, options: { bold: true, align: "center" } }, { text: "" },
    { text: `概要なしは約${total - 1.5}分`, options: { color: COL.gray } }]);
  s.addTable(rows, { x: 0.5, y: 1.4, w: 9.0, colW: [0.6, 3.4, 1.1, 1.6, 2.3], fontSize: 13, color: COL.text,
    border: { type: "solid", pt: 0.75, color: COL.line }, rowH: 0.42, valign: "middle", objectName: "時間配分表" });
  s.addNotes("【担当：司会／目安 0.5分】\n時間配分は目安。リハーサルで計測し、8〜12分に収まるよう調整すること。\n" +
    "2チーム目以降は「システム概要は省略します」と伝えて次へ進む。");

  // ===== 3. システム概要 =====
  pres.addSection({ title: "システム概要" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "システム概要" });
  s.addText("システム概要", { placeholder: "title" });
  tag(s, "担当：", owner.システム概要);
  sub(s, "総務部の社員が、従業員情報をWebブラウザから閲覧・登録・削除するシステム");
  const nodes = [["利用者", "総務部の社員"], ["Webブラウザ", "画面操作"], ["Webサーバ", "Tomcat\nServlet / JSP"],
    ["DBサーバ", "MySQL\nemp_sys_db"]];
  nodes.forEach((n, i) => {
    const x = 0.5 + i * 2.33;
    const first = i === 0;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: 1.85, h: 1.3, rectRadius: 0.1,
      fill: { color: first ? COL.teal : COL.pale }, line: { color: COL.teal, width: 1 }, objectName: "構成_" + n[0] });
    s.addText(n[0], { x, y: 1.6, w: 1.85, h: 0.4, fontSize: 15, bold: true, align: "center",
      color: first ? COL.white : COL.navy, margin: 0, isTextBox: true, objectName: "構成名_" + n[0] });
    s.addText(n[1], { x, y: 2.0, w: 1.85, h: 0.7, fontSize: 12, align: "center", valign: "top",
      color: first ? COL.white : COL.text, margin: 0, isTextBox: true, objectName: "構成説明_" + n[0] });
    if (i < nodes.length - 1) {
      s.addShape(pres.ShapeType.rightArrow, { x: x + 1.93, y: 2.0, w: 0.32, h: 0.3, fill: { color: COL.gray },
        line: { color: COL.gray }, objectName: "矢印" + (i + 1) });
    }
  });
  card(s, 0.5, 3.1, 4.4, 1.95, "目的・利用者", ["利用者：従業員情報を業務で扱う総務部の社員",
    "入退社のたびに発生する従業員情報の登録・削除を、ブラウザから行えるようにする",
    "ログインした利用者だけが操作できる"], COL.teal, "目的");
  card(s, 5.1, 3.1, 4.4, 1.95, "管理するデータ（3テーブル）", [
    "employee：従業員マスタ（コード・氏名・フリガナ・性別・生年月日・所属・入社日）",
    "section：所属マスタ（所属コード・部署名）", "user：ユーザマスタ（ユーザID・パスワード）"], COL.teal, "データ");
  s.addNotes("【目安 1.5分】※最初のチームのみ発表\n・誰が何のために使うシステムかを最初に伝える。\n" +
    "・構成図を指しながら「ブラウザ→Tomcat上のServlet/JSP→MySQL」の流れを説明する。");

  // ===== 4. 機能一覧 =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "システム概要" });
  s.addText("機能一覧と画面遷移", { placeholder: "title" });
  tag(s, "担当：", owner.システム概要);
  sub(s, hasExtra ? "必須機能5つ＋追加課題。ログイン後はメニュー画面から各機能へ遷移します"
    : "必須機能5つ。ログイン後はメニュー画面から各機能へ遷移します");
  const funcs = [
    ["ログイン", "ユーザID・パスワードで認証。成功→メニュー、失敗→エラー画面", "必須"],
    ["従業員一覧表示", "全従業員を表形式で表示（従業員コード昇順）", "必須"],
    ["従業員情報登録", "性別・所属はプルダウン、日付はカレンダーで入力", "必須"],
    ["従業員情報削除", "一覧で選択→確認ダイアログ→削除完了画面", "必須"],
    ["ログアウト", "セッションを破棄しログアウト画面を表示", "必須"],
  ];
  if (extra.更新) funcs.push(["従業員情報更新", "一覧で選択した従業員を初期表示して更新（コードは変更不可）", "追加"]);
  if (extra.ユーザ登録) funcs.push(["ユーザ情報登録", "システム利用者（ユーザID・パスワード）を登録", "追加"]);
  funcs.forEach((f, i) => {
    const y = 1.38 + i * 0.52;
    const color = f[2] === "必須" ? COL.teal : COL.gold;
    badge(s, 0.5, y + 0.04, i + 1, color, 0.38);
    s.addText(f[0], { x: 1.0, y, w: 2.0, h: 0.46, fontSize: 14, bold: true, color: COL.navy, valign: "middle",
      margin: 0, isTextBox: true, objectName: "機能名" + (i + 1) });
    s.addText(f[1], { x: 3.0, y, w: 5.5, h: 0.46, fontSize: 12, color: COL.text, valign: "middle",
      margin: 0, isTextBox: true, objectName: "機能説明" + (i + 1) });
    s.addText(f[2], { shape: pres.ShapeType.roundRect, rectRadius: 0.06, x: 8.65, y: y + 0.07, w: 0.85, h: 0.32,
      fontSize: 11, bold: true, align: "center", valign: "middle", color: COL.white, fill: { color },
      margin: 0, isTextBox: true, objectName: "区分" + (i + 1) });
  });
  s.addNotes("【目安 含む1.5分】\n・番号1〜5が必須機能、それ以降が追加課題。\n" +
    "・追加課題の表示は記入シートの「追加課題 > 実装した」で切り替わる。");

  // ===== 5. 開発環境 =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "システム概要" });
  s.addText("開発環境とプログラム構成", { placeholder: "title" });
  tag(s, "担当：", owner.システム概要);
  sub(s, "Servlet（受付）→ Logic（処理）→ DAO（DB操作）の3層に分けて実装");
  const envRows = data.開発環境.map((e) => [
    Object.assign(run(e.項目), { options: Object.assign({ bold: true, color: COL.navy, fill: { color: COL.pale } }, run(e.項目).options) }),
    run(e.内容),
  ]);
  s.addTable(envRows, { x: 0.5, y: 1.45, w: 3.6, colW: [1.1, 2.5], fontSize: 12, color: COL.text,
    border: { type: "solid", pt: 0.75, color: COL.line }, rowH: 0.42, valign: "middle", objectName: "開発環境表" });
  const layers = [
    ["servlet", "画面からの要求を受け付け、JSPへ振り分け", "Login / Logout / AdminMenu / EmployeeList / EmployeeCreate / EmployeeDelete / EmployeeEdit …"],
    ["model", "業務ロジックとデータの入れ物", "LoginLogic / EmployeeListLogic / EmployeeCreateLogic …\nEmployee / Section / User"],
    ["dao", "SQLを発行してDBを操作", "UserDAO / SectionDAO / EmployeeDAO"],
  ];
  layers.forEach((l, i) => {
    const y = 1.45 + i * 1.1;
    panel(s, 4.5, y, 5.0, 0.95, "層_" + l[0]);
    s.addText(l[0], { shape: pres.ShapeType.roundRect, rectRadius: 0.06, x: 4.62, y: y + 0.12, w: 1.0, h: 0.34,
      fill: { color: COL.teal }, color: COL.white, fontSize: 12, bold: true, align: "center", valign: "middle",
      margin: 0, isTextBox: true, objectName: "層名_" + l[0] });
    s.addText(l[1], { x: 5.75, y: y + 0.1, w: 3.65, h: 0.38, fontSize: 12, bold: true, color: COL.navy,
      valign: "middle", margin: 0, isTextBox: true, objectName: "層説明_" + l[0] });
    s.addText(l[2], { x: 4.62, y: y + 0.5, w: 4.8, h: 0.42, fontSize: 10, color: COL.gray,
      valign: "top", margin: 0, isTextBox: true, objectName: "層クラス_" + l[0] });
  });
  s.addNotes("【目安 含む1.5分】\n・クラス図の構成に沿って、servlet→model（Logic）→dao の順に呼び出される流れを伝える。");

  // ===== 6. 開発の進め方 =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "システム概要" });
  s.addText("開発の進め方（全30コマ）", { placeholder: "title" });
  tag(s, "担当：", owner.システム概要);
  sub(s, "要求分析から成果発表まで、WBSで進捗を管理しながら進めました");
  const phases = [["準備", "1〜2"], ["要求分析", "3〜5"], ["設計", "6〜10"], ["実装", "11〜20"],
    ["テスト", "21〜25"], ["納品", "26"], ["発表", "27〜30"]];
  const phaseColors = [COL.gray, COL.purple, COL.teal, COL.navy, COL.green, COL.gold, COL.orange];
  const pw = 9.0 / phases.length;
  phases.forEach((p, i) => {
    s.addText([{ text: p[0], options: { bold: true, fontSize: 13, breakLine: true } },
      { text: p[1] + "コマ", options: { fontSize: 10 } }], { shape: pres.ShapeType.homePlate, x: 0.5 + i * pw, y: 1.45,
      w: pw - 0.04, h: 0.9, fill: { color: phaseColors[i] }, line: { color: phaseColors[i] }, color: COL.white,
      align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "工程_" + p[0] });
  });
  card(s, 0.5, 2.85, 4.4, 2.2, "主な成果物", ["分析：ユースケース図・画面遷移図", "設計：シーケンス図・クラス図・画面設計書",
    "実装：ソースコード・DB作成SQL", "テスト：テスト仕様書兼結果報告書", "管理：WBS／進捗管理表"], COL.teal, "成果物");
  panel(s, 5.1, 2.85, 4.4, 2.2, "役割分担_背景");
  s.addText("役割分担", { x: 5.25, y: 2.95, w: 4.1, h: 0.35, fontSize: 14, bold: true, color: COL.teal,
    margin: 0, isTextBox: true, objectName: "役割分担_見出し" });
  const roleRuns = [];
  members.forEach((m, i) => {
    const n = memberNames[i];
    roleRuns.push({ text: n.text, options: Object.assign({ bold: true, color: COL.navy }, n.options) });
    roleRuns.push({ text: "：", options: { color: COL.text } });
    const r = run(m.担当作業);
    roleRuns.push({ text: r.text, options: Object.assign({ color: COL.text, breakLine: i < members.length - 1 }, r.options) });
  });
  s.addText(roleRuns, { x: 5.25, y: 3.35, w: 4.1, h: 1.6, fontSize: 13, valign: "top", margin: 0,
    paraSpaceAfter: 4, isTextBox: true, objectName: "役割分担_本文" });
  s.addNotes("【目安 含む1.5分】\n・工程の区切りはWBSの予定コマ。実績が大きくずれた工程があれば一言触れる。");

  // ===== 7. デモ：基本ルート =====
  pres.addSection({ title: "デモンストレーション" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "デモンストレーション" });
  s.addText("デモンストレーション：基本ルート", { placeholder: "title" });
  tag(s, "操作：", owner.デモ操作);
  sub(s, "Web画面を表示して、ログインからログアウトまでを順に操作します");
  const demo = [
    ["ログイン", "誤ったID／パスワード → エラー画面、正しい値 → メニュー画面（「ようこそ、○○さん」）"],
    ["登録", "メニュー →［従業員登録］→ 入力（性別・所属はプルダウン、日付はカレンダー）→ 登録完了画面"],
    ["一覧表示", "メニュー →［従業員一覧］→ 今登録した従業員が従業員コード順に表示されることを確認"],
    ["削除", "一覧で従業員を選択 →［削除］→ 確認ダイアログでOK → 削除完了画面 → 一覧から消えたことを確認"],
    ["ログアウト", "メニュー →［ログアウト］→ ログアウト画面 → ログイン画面へ戻る"],
  ];
  demo.forEach((d, i) => {
    const y = 1.38 + i * 0.66;
    badge(s, 0.5, y + 0.06, i + 1, COL.teal);
    s.addText(d[0], { x: 1.05, y, w: 1.45, h: 0.55, fontSize: 15, bold: true, color: COL.navy, valign: "middle",
      margin: 0, isTextBox: true, objectName: "デモ手順名" + (i + 1) });
    s.addText(d[1], { x: 2.5, y, w: 7.0, h: 0.55, fontSize: 12, color: COL.text, valign: "middle",
      margin: 0, isTextBox: true, objectName: "デモ手順説明" + (i + 1) });
  });
  const lid = run(data.ログインID, { color: COL.text });
  s.addText([
    { text: "URL：", options: { bold: true, color: COL.navy } },
    { text: "http://localhost:8080/empsys/　", options: { color: COL.text } },
    { text: "ログインID：", options: { bold: true, color: COL.navy } },
    lid,
  ], { x: 0.5, y: 4.75, w: 9.0, h: 0.35, fontSize: 12, margin: 0, isTextBox: true, objectName: "デモ環境" });
  s.addNotes("【目安 2.5分】\n・事前にTomcat・MySQLを起動し、ログイン画面を開いておく。\n" +
    "・まずエラー画面を見せてから正しくログインすると、入力チェックの説明になる。\n" +
    "・登録した従業員をそのまま削除すると、デモ後にデータが元に戻る。");

  // ===== 8. デモ：追加課題（実装した機能があるときだけ） =====
  if (hasExtra) {
    s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "デモンストレーション" });
    s.addText("デモンストレーション：追加課題", { placeholder: "title" });
    tag(s, "操作：", owner.デモ操作);
    sub(s, "要件定義書の「余力があれば」「追加システム要件」から実装した機能");
    const cards = [];
    if (extra.更新) cards.push(["従業員情報更新", ["一覧で従業員を選択 →［従業員更新］", "更新画面に選択した従業員の情報が初期表示される",
      "従業員コードは変更不可", "更新ボタン → 更新完了画面 → 一覧で変更を確認"]]);
    if (extra.ユーザ登録) cards.push(["ユーザ情報登録", ["メニュー →［システム利用者登録］", "ユーザID・パスワードを入力して登録",
      "登録完了画面を確認", "ログアウト後、登録したユーザでログインできることを確認"]]);
    const cw = cards.length === 2 ? 4.4 : 9.0;
    cards.forEach((c, i) => card(s, 0.5 + i * 4.6, 1.4, cw, 2.3, `${"①②"[i]} ${c[0]}`, c[1], COL.gold, "追加課題" + (i + 1)));
    s.addText("見どころ（ここを見てほしい点）", { x: 0.5, y: 3.9, w: 9.0, h: 0.3, fontSize: 13, bold: true,
      color: COL.teal, margin: 0, isTextBox: true, objectName: "見どころ_見出し" });
    const hl = run(extra.見どころ, { color: COL.text });
    s.addText(hl.text, Object.assign({ x: 0.5, y: 4.25, w: 9.0, h: 0.85, fontSize: 14, valign: "top", margin: 0.12,
      fill: { color: hl.options.italic ? COL.fillBg : COL.pale },
      line: hl.options.italic ? { color: COL.orange, width: 1, dashType: "dash" } : { color: COL.pale },
      isTextBox: true, objectName: "見どころ_本文" }, hl.options));
    s.addNotes("【目安 1.5分】\n・独自に追加した機能（入力チェックの強化、画面デザインなど）があれば、ここで見せる。");
  }

  // ===== 9. 学んだこと =====
  pres.addSection({ title: "振り返り" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "振り返り" });
  s.addText("学んだこと", { placeholder: "title" });
  tag(s, "担当：", owner.学んだこと);
  sub(s, "要求分析からテスト・納品まで、ひと通りの開発工程を経験して学んだこと");
  const ln = data.学んだこと.length;
  const lw = (9.0 - 0.21 * (ln - 1)) / ln;
  data.学んだこと.forEach((l, i) => {
    const x = 0.5 + i * (lw + 0.21);
    panel(s, x, 1.45, lw, 3.55, "学び背景" + (i + 1));
    badge(s, x + 0.2, 1.62, i + 1, COL.teal);
    txt(s, l.観点, { x: x + 0.75, y: 1.62, w: lw - 0.9, h: 0.42, fontSize: 15, bold: true, color: COL.navy,
      valign: "middle", objectName: "学び見出し" + (i + 1) });
    txt(s, l.内容, { x: x + 0.2, y: 2.25, w: lw - 0.4, h: 2.6, fontSize: 13, valign: "top", color: COL.text,
      objectName: "学び内容" + (i + 1) });
  });
  s.addNotes("【目安 1分】\n・「何を学んだか」に加えて「どの場面でそう感じたか」を一言添えると伝わりやすい。");

  // ===== 10. 苦労したこと =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "振り返り" });
  s.addText("苦労したこと・難しかったこと", { placeholder: "title" });
  tag(s, "担当：", owner.苦労したこと);
  sub(s, "「何が起きたか → どう解決したか」の形でまとめました");
  s.addText("何が起きたか", { x: 2.9, y: 1.4, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: COL.gray,
    margin: 0, isTextBox: true, objectName: "列見出し_問題" });
  s.addText("どう解決したか", { x: 6.3, y: 1.4, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: COL.gray,
    margin: 0, isTextBox: true, objectName: "列見出し_解決" });
  data.苦労したこと.forEach((h, i) => {
    const y = 1.78 + i * 1.12;
    panel(s, 0.5, y, 9.0, 1.0, "苦労背景" + (i + 1));
    badge(s, 0.65, y + 0.29, i + 1, COL.orange);
    txt(s, h.見出し, { x: 1.2, y: y + 0.08, w: 1.6, h: 0.84, fontSize: 13, bold: true, color: COL.navy,
      valign: "middle", objectName: "苦労見出し" + (i + 1) });
    txt(s, h.何が起きたか, { x: 2.9, y: y + 0.08, w: 3.2, h: 0.84, fontSize: 11, color: COL.text,
      valign: "middle", objectName: "苦労内容" + (i + 1) });
    s.addShape(pres.ShapeType.rightArrow, { x: 6.12, y: y + 0.38, w: 0.16, h: 0.24, fill: { color: COL.gray },
      line: { color: COL.gray }, objectName: "苦労矢印" + (i + 1) });
    txt(s, h.どう解決したか, { x: 6.3, y: y + 0.08, w: 3.1, h: 0.84, fontSize: 11, color: COL.text,
      valign: "middle", objectName: "苦労解決" + (i + 1) });
  });
  s.addNotes("【目安 1分】\n・チーム運営面（分担、スケジュール、マージ時の衝突など）の苦労があれば1件入れるとバランスがよい。");

  // ===== 11. 工夫点 =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "振り返り" });
  s.addText("工夫点（アピールしたい点）", { placeholder: "title" });
  tag(s, "担当：", owner.工夫点);
  sub(s, "左はテスト結果から言える数字、右はチームで特に見てほしい点");
  const t = data.テスト結果;
  const stats = [[t.テスト項目数, "テスト項目"], [t.OK, "OK"], [t.NG, "NG"]];
  stats.forEach((st, i) => {
    const y = 1.4 + i * 1.12;
    const dark = i === 2;
    s.addShape(pres.ShapeType.roundRect, { x: 0.5, y, w: 3.0, h: 1.0, rectRadius: 0.08,
      fill: { color: dark ? COL.teal : COL.pale }, line: { color: dark ? COL.teal : COL.pale }, objectName: "数値背景" + (i + 1) });
    txt(s, st[0], { x: 0.6, y: y + 0.06, w: 1.6, h: 0.88, fontSize: 36, bold: true, align: "center", valign: "middle",
      color: dark ? COL.white : COL.teal, fillColor: dark ? COL.fillOnDark : COL.orange, objectName: "数値" + (i + 1) }, "—");
    s.addText(st[1], { x: 2.2, y: y + 0.06, w: 1.2, h: 0.88, fontSize: 14, bold: true, valign: "middle",
      color: dark ? COL.white : COL.navy, margin: 0, isTextBox: true, objectName: "数値ラベル" + (i + 1) });
  });
  txt(s, t.注記, { x: 0.5, y: 4.78, w: 3.2, h: 0.22, fontSize: 9, color: COL.gray, objectName: "数値注記" }, "");
  data.工夫点.forEach((a, i) => {
    const y = 1.4 + i * 1.22;
    badge(s, 3.9, y + 0.05, i + 1, COL.teal);
    txt(s, a.見出し, { x: 4.45, y, w: 5.05, h: 0.45, fontSize: 15, bold: true, valign: "middle", color: COL.navy,
      objectName: "工夫見出し" + (i + 1) });
    txt(s, a.説明, { x: 4.45, y: y + 0.45, w: 5.05, h: 0.65, fontSize: 12, valign: "top", color: COL.text,
      objectName: "工夫説明" + (i + 1) });
  });
  s.addNotes("【目安 1分】\n・数字はテスト仕様書兼結果報告書の最終版に合わせる。\n・チーム独自の工夫は、デモで見せた画面と結びつけると印象に残る。");

  // ===== 12. 個人ごとの感想 =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "振り返り" });
  s.addText("個人ごとの感想", { placeholder: "title" });
  s.addText("全員／1人30秒", { shape: pres.ShapeType.roundRect, rectRadius: 0.08, x: 7.4, y: 0.38, w: 2.1, h: 0.42,
    fontSize: 11, bold: true, color: COL.orange, fill: { color: COL.tagBg }, line: { color: COL.orange, width: 1 },
    align: "center", valign: "middle", isTextBox: true, objectName: "担当タグ" });
  sub(s, "担当した作業と、良かった点・反省点・次に活かしたいこと");
  const four = members.length === 4;
  members.forEach((m, i) => {
    // 4人は2×2、3人は横に3列
    const w = four ? 4.4 : 2.86;
    const h = four ? 1.7 : 3.55;
    const x = four ? 0.5 + (i % 2) * 4.6 : 0.5 + i * 3.07;
    const y = four ? 1.4 + Math.floor(i / 2) * 1.85 : 1.45;
    panel(s, x, y, w, h, "感想背景" + (i + 1));
    badge(s, x + 0.15, y + 0.15, i + 1, COL.teal, 0.38);
    const n = memberNames[i];
    s.addText([{ text: n.text, options: Object.assign({ bold: true, color: COL.navy }, n.options) }], {
      x: x + 0.65, y: y + 0.13, w: w - 0.8, h: 0.42, fontSize: 14, valign: "middle", margin: 0, isTextBox: true,
      objectName: "感想氏名" + (i + 1) });
    const items = [["良かった点", m.良かった点], ["反省点", m.反省点], ["次に活かしたいこと", m.次に活かしたいこと]];
    const runs = [];
    items.forEach(([label, v], j) => {
      const r = run(v, { color: COL.text });
      runs.push({ text: label + "：", options: { bold: true, color: COL.navy, bullet: true } });
      runs.push({ text: r.text, options: Object.assign({ breakLine: j < items.length - 1 }, r.options) });
    });
    s.addText(runs, { x: x + 0.2, y: y + 0.62, w: w - 0.4, h: h - 0.72, fontSize: 12, valign: "top", margin: 0,
      paraSpaceAfter: 2, isTextBox: true, objectName: "感想本文" + (i + 1) });
  });
  s.addNotes("【担当：全員／目安 2分（1人30秒）】\n・スライドには要点だけ書き、詳しくは口頭で話す。");

  // ===== 13. 今後の授業への要望 =====
  pres.addSection({ title: "要望・まとめ" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "要望・まとめ" });
  s.addText("今後の授業への要望", { placeholder: "title" });
  tag(s, "担当：", owner.要望);
  sub(s, "企業課題Iに向けて、そして来年の3年生の授業に向けて");
  [["企業課題Iに向けた要望", data.企業課題Iへの要望], ["来年の3年生授業への改善点", data.来年の授業への改善点]].forEach((col, i) => {
    const x = 0.5 + i * 4.6;
    panel(s, x, 1.4, 4.4, 3.65, "要望背景" + (i + 1));
    s.addText(col[0], { x: x + 0.2, y: 1.55, w: 4.0, h: 0.4, fontSize: 15, bold: true, color: COL.teal,
      margin: 0, isTextBox: true, objectName: "要望見出し" + (i + 1) });
    s.addText(bulletRuns(col[1], { color: COL.text }), { x: x + 0.2, y: 2.05, w: 4.0, h: 2.85, fontSize: 13,
      valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "要望本文" + (i + 1) });
  });
  s.addNotes("【目安 1分】\n・不満ではなく「こうなると良い」という提案の形で伝える。");

  // ===== 14. まとめ =====
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "要望・まとめ" });
  s.addText("ご清聴ありがとうございました", { placeholder: "title" });
  const endRuns = [onDark(run(data.チーム名)), { text: "　", options: {} }];
  memberNames.forEach((r, i) => {
    if (i > 0) endRuns.push({ text: "／", options: { color: COL.white } });
    endRuns.push(onDark(r));
  });
  s.addText(endRuns, { x: 0.6, y: 3.0, w: 8.8, h: 0.4, fontSize: 16, bold: true, margin: 0, isTextBox: true, objectName: "結びメンバー" });
  s.addText("ご質問があればお願いします", { x: 0.6, y: 3.5, w: 8.8, h: 0.4, fontSize: 14, color: COL.lightText,
    margin: 0, isTextBox: true, objectName: "質疑" });
  s.addNotes("【担当：リーダー／目安 0.5分】\n全員で礼をして締める。");

  return pres;
}

async function main(input = DEFAULT_INPUT, output = DEFAULT_OUTPUT) {
  const data = parse(fs.readFileSync(input, "utf8"));
  const pres = build(data);
  await pres.writeFile({ fileName: output });
  return output;
}

if (require.main === module) {
  main(process.argv[2], process.argv[3])
    .then((out) => console.log("作成しました: " + out))
    .catch((e) => {
      console.error(e.message);
      process.exit(1);
    });
}

module.exports = { build, main };
