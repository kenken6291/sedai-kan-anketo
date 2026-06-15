// ════════════════════════════════════════════════════════════════
//  世代間アンケート — Google Apps Script
//  ※ GASはブラウザからのPOSTのCORSプリフライトを通過できないため、
//     すべての操作をGETパラメータで受け取る設計にしています。
// ════════════════════════════════════════════════════════════════

const SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');
const SH_QUESTIONS = 'Questions';
const SH_VOTES     = 'Votes';

// ────────────────────────────────────────────────────────────────
//  唯一のエントリポイント：すべてdoGetで処理
// ────────────────────────────────────────────────────────────────
function doGet(e) {
  const p      = e.parameter;
  const action = p.action || '';
  try {
    if (action === 'getQuestions')   return ok(getQuestions());
    if (action === 'getResults')     return ok(getResults());
    if (action === 'saveQuestion')   return ok(saveQuestion(p));
    if (action === 'deleteQuestion') return ok(deleteQuestion(p));
    if (action === 'castVote')       return ok(castVote(p));
    return err('Unknown action: ' + action);
  } catch(ex) {
    return err(ex.message);
  }
}

// ────────────────────────────────────────────────────────────────
//  シート初期化（初回だけ手動実行）
// ────────────────────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SS_ID);

  let shQ = ss.getSheetByName(SH_QUESTIONS);
  if (!shQ) {
    shQ = ss.insertSheet(SH_QUESTIONS);
    shQ.appendRow([
      'id','title','badge','authorGen',
      'opt0_emoji','opt0_text',
      'opt1_emoji','opt1_text',
      'opt2_emoji','opt2_text',
      'opt3_emoji','opt3_text',
      'createdAt'
    ]);
    shQ.setFrozenRows(1);
    shQ.getRange(1,1,1,13).setFontWeight('bold');
    shQ.appendRow([
      'q_date','デートの費用は「割り勘」にするべき？','FALSE',0,
      '🤝','割り勘にすべき',
      '🎩','年長者・誘った側が多めに払うべき',
      '','','','',new Date().toISOString()
    ]);
    shQ.appendRow([
      'q_wedding','結婚式の規模・スタイルについてどう思う？','TRUE',2,
      '⛪','伝統的な式・披露宴をやりたい',
      '👨‍👩‍👧','身内だけの少人数でやりたい',
      '📸','フォトウェディングだけでいい',
      '✌️','式は不要・ナシ婚でいい',
      new Date().toISOString()
    ]);
  }

  let shV = ss.getSheetByName(SH_VOTES);
  if (!shV) {
    shV = ss.insertSheet(SH_VOTES);
    shV.appendRow(['questionId','genIndex','optionIndex','timestamp']);
    shV.setFrozenRows(1);
    shV.getRange(1,1,1,4).setFontWeight('bold');
  }

  return 'Sheets initialized.';
}

// ────────────────────────────────────────────────────────────────
//  質問一覧取得
// ────────────────────────────────────────────────────────────────
function getQuestions() {
  const rows = getSheet(SH_QUESTIONS).getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(rowToQuestion).filter(q => q.id);
}

// ────────────────────────────────────────────────────────────────
//  集計取得
// ────────────────────────────────────────────────────────────────
function getResults() {
  const questions = getQuestions();
  const rows = getSheet(SH_VOTES).getDataRange().getValues().slice(1);
  const tally = {};
  questions.forEach(q => {
    tally[q.id] = [0,1,2,3].map(() => q.options.map(() => 0));
  });
  rows.forEach(r => {
    const [qid, gi, oi] = [String(r[0]), Number(r[1]), Number(r[2])];
    if (tally[qid] && tally[qid][gi] && tally[qid][gi][oi] !== undefined) {
      tally[qid][gi][oi]++;
    }
  });
  return tally;
}

// ────────────────────────────────────────────────────────────────
//  質問保存（GETパラメータで受け取り）
//  options は JSON.stringify した文字列を "options" パラメータで渡す
// ────────────────────────────────────────────────────────────────
function saveQuestion(p) {
  const id        = p.id;
  const title     = p.title;
  const badge     = p.badge === 'true';
  const authorGen = p.authorGen !== '' && p.authorGen != null ? Number(p.authorGen) : null;
  const options   = JSON.parse(p.options); // [{emoji,text}, ...]

  const sh   = getSheet(SH_QUESTIONS);
  const data = sh.getDataRange().getValues();
  const row  = buildRow(id, title, badge, authorGen, options);
  const idx  = data.findIndex((r, i) => i > 0 && String(r[0]) === id);

  if (idx > 0) {
    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
    return { status:'updated', id };
  } else {
    sh.appendRow(row);
    return { status:'created', id };
  }
}

// ────────────────────────────────────────────────────────────────
//  質問削除
// ────────────────────────────────────────────────────────────────
function deleteQuestion(p) {
  const id   = p.id;
  const sh   = getSheet(SH_QUESTIONS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sh.deleteRow(i + 1);
      deleteVotesByQuestion(id);
      return { status:'deleted', id };
    }
  }
  throw new Error('Question not found: ' + id);
}

// ────────────────────────────────────────────────────────────────
//  投票記録
// ────────────────────────────────────────────────────────────────
function castVote(p) {
  getSheet(SH_VOTES).appendRow([
    p.questionId,
    Number(p.genIndex),
    Number(p.optionIndex),
    new Date().toISOString()
  ]);
  return { status:'voted' };
}

// ────────────────────────────────────────────────────────────────
//  ヘルパー
// ────────────────────────────────────────────────────────────────
function getSheet(name) {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function rowToQuestion(r) {
  const opts = [];
  for (let i = 4; i <= 10; i += 2) {
    if (String(r[i+1]).trim() !== '') {
      opts.push({ emoji: String(r[i]) || '▶', text: String(r[i+1]) });
    }
  }
  return {
    id:        String(r[0]),
    title:     String(r[1]),
    badge:     String(r[2]).toUpperCase() === 'TRUE',
    authorGen: (r[3] !== '' && r[3] != null) ? Number(r[3]) : null,
    options:   opts,
  };
}

function buildRow(id, title, badge, authorGen, options) {
  const opts = options.slice(0, 4);
  const row  = [id, title, badge ? 'TRUE' : 'FALSE', authorGen ?? ''];
  for (let i = 0; i < 4; i++) {
    row.push(opts[i]?.emoji ?? '');
    row.push(opts[i]?.text  ?? '');
  }
  row.push(new Date().toISOString());
  return row;
}

function deleteVotesByQuestion(qid) {
  const sh   = getSheet(SH_VOTES);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === qid) sh.deleteRow(i + 1);
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok:true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok:false, error:msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
