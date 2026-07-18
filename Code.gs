/**
 * 実験1・実験2 自動化ウェブシステム — バックエンド (Google Apps Script)
 *
 * 前提：このスクリプトは「コンテナバインド型」として、Sheetテンプレート
 * （実験1_2_システムSheetテンプレート.xlsx をGoogleスプレッドシートに変換したもの）
 * の [拡張機能] > [Apps Script] から作成することを想定している。
 * そのため SpreadsheetApp.getActiveSpreadsheet() で対象シートが自動的に取れる。
 *
 * シート構成（4タブ、ヘッダーは1行目固定）：
 *   Codes            : code | used | used_at | session_id | exclude_material | expected_name | email
 *   Consent          : session_id | subject_id | name | consent_date | timestamp | expected_name | email | name_match
 *   Responses        : session_id | subject_id | status | started_at | age | gender |
 *                       native_language | vviq_json | pre_saved_at | material_id |
 *                       condition | assign_saved_at | audio_start_at | audio_end_at |
 *                       audio_compliance_flags | post_json | post_saved_at
 *   AssignmentPool   : material_id | condition | target_n | current_n
 *
 * 名簿連携（任意機能）：Codesシートに expected_name / email 列を追加しておくと、
 * 学内募集などで事前に参加者が判明している場合、その人に配布する口令へ
 * あらかじめ氏名・メールアドレスを紐付けておくことができる。
 * 同意画面で入力された氏名は、この expected_name と突き合わせる（全角/半角・空白の表記ゆれは
 * 吸収するが、それ以外の不一致は参加をブロックする＝申込み本人と実験参加者の同一性を保証するため）。
 * Codesシートにemailも設定されている場合は、氏名一致後にメールアドレスも突き合わせ、
 * 不一致なら「お申込み時のメールアドレスを入力してください」という専用メッセージでブロックする
 * （1人が複数メールを使い分けているケースを考慮した文言）。ブロック時は研究者の連絡先
 * （RESEARCHER_CONTACT_EMAIL）を案内する。この2列を空欄のままにしておけば、
 * 従来通り照合なしの匿名口令として機能する（テスト用コード等）。
 */

// ============ 音声材料のURLをここに設定する（Google Driveの共有リンクのfile IDなど） ============
// key は `${materialId}_${condition}` の形式。condition は "高群" / "低群"。
// GitHub（raw.githubusercontent.com）でホスティング。Google DriveのuncリンクはORB
// （Chromeのクロスオリジン保護）でブロックされるため使用不可と判明したため切り替えた。
// リポジトリ名・ユーザー名・ブランチ名が異なる場合はここを書き換えること。
const MATERIAL_AUDIO_URLS = {
  "M1_高群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m1_high.mp3",
  "M1_低群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m1_low.mp3",
  "M2_高群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m2_high.mp3",
  "M2_低群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m2_low.mp3",
  "M3_高群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m3_high.mp3",
  "M3_低群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m3_low.mp3",
  "M4_高群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m4_high.mp3",
  "M4_低群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m4_low.mp3",
  "M5_高群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m5_high.mp3",
  "M5_低群": "https://raw.githubusercontent.com/dsreameng/audio-materials/main/m5_low.mp3",
};

const SHEET = {
  CODES: "Codes",
  CONSENT: "Consent",
  RESPONSES: "Responses",
  POOL: "AssignmentPool",
};

// 名簿連携で氏名・メールアドレスの不一致が起きた場合に、参加者へ案内する問い合わせ先。
// 必要に応じて実際の連絡先メールアドレスに書き換えること。
const RESEARCHER_CONTACT_EMAIL = "dsreameng@gmail.com";

// ============ エントリポイント ============
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("実験参加")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============ 共通ヘルパー ============
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("シートが見つかりません: " + name);
  return sh;
}

function headerIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h] = i + 1; }); // 1-indexed column
  return map;
}

function findRowByValue_(sheet, colIndex, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2; // actual sheet row
  }
  return -1;
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
}

// 実験ID（P001, P002...）の原子的な発番。スクリプトプロパティにカウンタを保持する。
function nextSubjectId_() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let n = Number(props.getProperty("SUBJECT_COUNTER") || "0");
    n += 1;
    props.setProperty("SUBJECT_COUNTER", String(n));
    return "P" + ("000" + n).slice(-3);
  } finally {
    lock.releaseLock();
  }
}

// ============ ① 口令の検証（まだ消費しない） ============
// 方針：口令が実際に「使用済み」になるのは、参加者が説明・同意・事前アンケート・環境確認を終え、
// 画面4で「次へ（音声材料へ進む）」を押した瞬間（＝commitCode実行時）。
// それより前（説明を読んだだけ、環境が整わず離脱した等）の段階では口令を消費しないため、
// 同じコードをそのまま使い直せる。ここではコードの存在確認とセッション（Responses行）の
// 仮作成のみを行う。
function validateCode(code) {
  code = (code || "").trim();
  if (!code) return { success: false, reason: "口令が入力されていません。" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_(SHEET.CODES);
    const h = headerIndex_(sh);
    const row = findRowByValue_(sh, h["code"], code);
    if (row === -1) {
      return { success: false, reason: "口令が正しくありません。" };
    }
    const used = sh.getRange(row, h["used"]).getValue();
    if (used === true || String(used).toUpperCase() === "TRUE") {
      return { success: false, reason: "この口令は既に使用されています。再度の参加はできません。" };
    }

    // Responsesシートに仮のレコードを作成する（口令自体はまだ未消費のまま）
    const sessionId = Utilities.getUuid();
    const rsh = getSheet_(SHEET.RESPONSES);
    rsh.appendRow([sessionId, "", "started", nowStr_()]);

    return { success: true, sessionId: sessionId };
  } finally {
    lock.releaseLock();
  }
}

// ============ ①' 口令の正式消費（「次へ（音声材料へ進む）」を押した時点で呼ばれる） ============
// Codesシートの exclude_material 列（任意）：例 "M3_低群"。
// 途中離脱者への再発行コードにこれを設定しておくと、その参加者が新しいコードで
// 再度参加した際、以前割り当てられていたのと同じ材料×条件には二度と割り当てない
// （＝同じ内容を聴き直してしまう事態を避ける）。通常のコードでは空欄のままでよい。
function commitCode(sessionId, code) {
  code = (code || "").trim();
  if (!code) return { success: false, reason: "口令が不正です。" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_(SHEET.CODES);
    const h = headerIndex_(sh);
    const row = findRowByValue_(sh, h["code"], code);
    if (row === -1) {
      return { success: false, reason: "口令が見つかりません。実験者にご連絡ください。" };
    }
    const used = sh.getRange(row, h["used"]).getValue();
    if (used === true || String(used).toUpperCase() === "TRUE") {
      return { success: false, reason: "この口令は既に使用されています。" };
    }
    const excludeMaterial = h["exclude_material"] ? String(sh.getRange(row, h["exclude_material"]).getValue() || "").trim() : "";

    // ここで初めて使用済みにマークする（この時点から二重利用・再開を完全に防止する）
    sh.getRange(row, h["used"]).setValue(true);
    sh.getRange(row, h["used_at"]).setValue(nowStr_());
    sh.getRange(row, h["session_id"]).setValue(sessionId);

    if (excludeMaterial) {
      const rsh = getSheet_(SHEET.RESPONSES);
      const rh = headerIndex_(rsh);
      const sessionRow = findRowByValue_(rsh, rh["session_id"], sessionId);
      if (sessionRow !== -1 && rh["exclude_material"]) {
        rsh.getRange(sessionRow, rh["exclude_material"]).setValue(excludeMaterial);
      }
    }

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 氏名・メールアドレスの表記ゆれ（全角/半角、前後・途中の空白）を吸収するための正規化。
// NFKCで全角英数字・記号を半角に統一したうえで、あらゆる空白（全角スペース含む）を除去する。
// 漢字そのものの表記（例：「渡辺」と「渡邊」）までは吸収しないため、完全な別人・別表記の場合は
// 依然として不一致として検出される（＝下のブロック処理の対象になる）。
function normalizeName_(s) {
  return String(s || "").normalize("NFKC").replace(/[\s　]/g, "").trim();
}
function normalizeEmail_(s) {
  return String(s || "").normalize("NFKC").replace(/[\s　]/g, "").trim().toLowerCase();
}

// ============ ② 同意の記録・実験IDの発番 ============
// code: 任意（省略可、後方互換のため）。指定された場合、Codesシートの expected_name / email 列と
// 入力された氏名・メールアドレスを突き合わせる。
//   - 氏名が一致しない場合：参加をブロックする（申込み本人と実験参加者が一致することを保証するため。
//     謝礼支給・倫理審査上、別人が入力した場合に進めてしまうことは避ける必要がある）。
//   - 氏名が一致し、かつCodesシートにemailが登録されている場合：メールアドレスも突き合わせる。
//     一致しない場合もブロックするが、「申込み時に使ったメールアドレスを入力してください」という
//     専用の案内を出す（1人が複数のメールアドレスを使い分けているケースを考慮し、迷わず入力し直せるように）。
//   - Codesシートに expected_name が設定されていないコード（匿名コード・テスト用コード等）は、
//     従来通り照合なしでそのまま通す。
function recordConsent(sessionId, name, consentDate, code, email) {
  if (!sessionId) return { success: false, reason: "セッションが不正です。" };
  name = (name || "").trim();
  email = (email || "").trim();
  if (!name) return { success: false, reason: "氏名を入力してください。" };

  // 名簿照合（Codesシートに expected_name 列があり、かつ該当コードにその値が入っている場合のみ）
  let expectedName = "";
  let expectedEmail = "";
  if (code) {
    const csh_codes = getSheet_(SHEET.CODES);
    const ch = headerIndex_(csh_codes);
    if (ch["expected_name"]) {
      const codeRow = findRowByValue_(csh_codes, ch["code"], (code || "").trim());
      if (codeRow !== -1) {
        expectedName = String(csh_codes.getRange(codeRow, ch["expected_name"]).getValue() || "").trim();
        expectedEmail = ch["email"] ? String(csh_codes.getRange(codeRow, ch["email"]).getValue() || "").trim() : "";
      }
    }
  }

  const nameMatch = expectedName ? (normalizeName_(expectedName) === normalizeName_(name)) : "";
  if (expectedName && !nameMatch) {
    return {
      success: false,
      reason: "お名前が、お申込み時にご登録いただいたお名前と一致しません。お手数ですが、お申込みの際にご使用になったお名前でご入力ください。解決しない場合は " + RESEARCHER_CONTACT_EMAIL + " までご連絡ください。",
    };
  }
  let emailMatch = "";
  if (expectedName && expectedEmail) {
    emailMatch = normalizeEmail_(expectedEmail) === normalizeEmail_(email);
    if (!emailMatch) {
      return {
        success: false,
        reason: "メールアドレスが、お申込み時にご登録いただいたものと一致しません。お申込みの際にご使用になったメールアドレスをご入力ください。解決しない場合は " + RESEARCHER_CONTACT_EMAIL + " までご連絡ください。",
      };
    }
  }

  const subjectId = nextSubjectId_();

  // 同意記録（氏名・メールアドレスを含む）は別シートに保存
  const csh = getSheet_(SHEET.CONSENT);
  const consentHeader = headerIndex_(csh);
  const newRow = new Array(csh.getLastColumn()).fill("");
  newRow[(consentHeader["session_id"] || 1) - 1] = sessionId;
  newRow[(consentHeader["subject_id"] || 2) - 1] = subjectId;
  newRow[(consentHeader["name"] || 3) - 1] = name;
  newRow[(consentHeader["consent_date"] || 4) - 1] = consentDate;
  newRow[(consentHeader["timestamp"] || 5) - 1] = nowStr_();
  if (consentHeader["expected_name"]) newRow[consentHeader["expected_name"] - 1] = expectedName;
  if (consentHeader["email"]) newRow[consentHeader["email"] - 1] = email || expectedEmail;
  if (consentHeader["name_match"]) newRow[consentHeader["name_match"] - 1] = nameMatch === "" ? "" : nameMatch;
  csh.appendRow(newRow);

  // Responses側にはsubject_idのみ書き込む（氏名は書かない）
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false, reason: "セッションが見つかりません。" };
  rsh.getRange(row, h["subject_id"]).setValue(subjectId);
  rsh.getRange(row, h["status"]).setValue("consented");

  return { success: true, subjectId: subjectId };
}

// ============ ③ 事前アンケートの保存 ============
// payload 例: { instructedCheck: 3, vviq: {...}, mood: 5 }
// 年齢・性別・母語はこの段階では聞かない（実験1_事前アンケートには含まれず、聴取後アンケートのD.セクションで収集する）。
function savePreQuestionnaire(sessionId, payload) {
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false, reason: "セッションが見つかりません。" };

  rsh.getRange(row, h["vviq_json"]).setValue(JSON.stringify(payload || {}));
  rsh.getRange(row, h["pre_saved_at"]).setValue(nowStr_());
  rsh.getRange(row, h["status"]).setValue("pre_done");

  return { success: true };
}

// ============ ④ 環境・ヘッドホン確認の保存 ============
// envQuiet: 参加者の自己申告（静かな環境か）。マイクによる客観測定は、Apps Script配信ページが
// iframeでラップされマイク権限が許可されないため断念し、自己申告に戻した（env_noise_level列に
// TRUE/FALSEとして格納する）。
function saveEnvCheck(sessionId, envQuiet, headphoneConfirmed, envCheckMethod, envNoiseValue) {
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false, reason: "セッションが見つかりません。" };

  rsh.getRange(row, h["env_noise_level"]).setValue(!!envQuiet);
  rsh.getRange(row, h["headphone_confirmed"]).setValue(!!headphoneConfirmed);
  rsh.getRange(row, h["env_saved_at"]).setValue(nowStr_());
  rsh.getRange(row, h["status"]).setValue("env_checked");

  // 任意列（Responsesシートに無ければ何もしない＝後方互換。追加したい場合はヘッダーに
  // env_check_method（"mic"/"self_report"）と env_noise_value（マイク実測時の相対値）を足す）
  if (h["env_check_method"]) {
    rsh.getRange(row, h["env_check_method"]).setValue(envCheckMethod || "self_report");
  }
  if (h["env_noise_value"] && envNoiseValue !== undefined && envNoiseValue !== null) {
    rsh.getRange(row, h["env_noise_value"]).setValue(envNoiseValue);
  }

  return { success: true };
}

// ============ ⑤ 刺激材料のランダム割当（人数バランス・排他制御あり） ============
function assignMaterial(sessionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // このセッションに割当除外の指定があるか確認する（途中離脱者への再発行コード用）
    const rshCheck = getSheet_(SHEET.RESPONSES);
    const rhCheck = headerIndex_(rshCheck);
    const sessionRow = findRowByValue_(rshCheck, rhCheck["session_id"], sessionId);
    let excludeMaterial = "";
    if (sessionRow !== -1 && rhCheck["exclude_material"]) {
      excludeMaterial = String(rshCheck.getRange(sessionRow, rhCheck["exclude_material"]).getValue() || "").trim();
    }

    const psh = getSheet_(SHEET.POOL);
    const ph = headerIndex_(psh);
    const lastRow = psh.getLastRow();
    const rows = psh.getRange(2, 1, lastRow - 1, psh.getLastColumn()).getValues();

    // target_n に対して current_n が最も余裕のある行を候補にする（除外指定があればまず除く）
    function pickCandidates(applyExclusion) {
      let cands = [];
      let maxSlack = -Infinity;
      rows.forEach((r, i) => {
        const materialId = r[ph["material_id"] - 1];
        const condition = r[ph["condition"] - 1];
        if (applyExclusion && excludeMaterial && (materialId + "_" + condition) === excludeMaterial) return;
        const target = Number(r[ph["target_n"] - 1]);
        const current = Number(r[ph["current_n"] - 1]);
        const slack = target - current;
        if (slack > maxSlack) {
          maxSlack = slack;
          cands = [i];
        } else if (slack === maxSlack) {
          cands.push(i);
        }
      });
      return cands;
    }

    let candidates = pickCandidates(true);
    if (candidates.length === 0) {
      // 除外を適用すると候補がなくなる場合（残り全セルが除外対象のみ等）はやむを得ず除外を解除する
      candidates = pickCandidates(false);
    }

    if (candidates.length === 0) {
      return { success: false, reason: "割当プールが空です。実験者に連絡してください。" };
    }

    const pick = candidates[Math.floor(Math.random() * candidates.length)];
    const rowNum = pick + 2;
    const materialId = rows[pick][ph["material_id"] - 1];
    const condition = rows[pick][ph["condition"] - 1];
    const currentN = Number(rows[pick][ph["current_n"] - 1]);
    psh.getRange(rowNum, ph["current_n"]).setValue(currentN + 1);

    // Responsesに反映
    const rsh = getSheet_(SHEET.RESPONSES);
    const h = headerIndex_(rsh);
    const row = findRowByValue_(rsh, h["session_id"], sessionId);
    if (row === -1) return { success: false, reason: "セッションが見つかりません。" };
    rsh.getRange(row, h["material_id"]).setValue(materialId);
    rsh.getRange(row, h["condition"]).setValue(condition);
    rsh.getRange(row, h["assign_saved_at"]).setValue(nowStr_());
    rsh.getRange(row, h["status"]).setValue("material_assigned");

    const key = materialId + "_" + condition;
    const audioUrl = MATERIAL_AUDIO_URLS[key] || "";

    return { success: true, materialId: materialId, condition: condition, audioUrl: audioUrl };
  } finally {
    lock.releaseLock();
  }
}

// ============ ⑤ 音声再生の開始・終了記録 ============
function recordAudioStart(sessionId) {
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false };
  rsh.getRange(row, h["audio_start_at"]).setValue(nowStr_());
  return { success: true };
}

// complianceFlagsに networkTerminated / voluntaryWithdrawal が立っている場合、statusにその旨を残す。
// これにより Responses シートの status 列を見るだけで「正常に聴取完了」「通信不良で強制終了」
// 「本人が自発的に中止」を区別でき、口令再発行が必要なケースをすぐ見つけられる。
function recordAudioEnd(sessionId, complianceFlags) {
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false };
  rsh.getRange(row, h["audio_end_at"]).setValue(nowStr_());
  rsh.getRange(row, h["audio_compliance_flags"]).setValue(JSON.stringify(complianceFlags || {}));
  let status = "audio_done";
  if (complianceFlags && complianceFlags.networkTerminated) status = "network_terminated";
  else if (complianceFlags && complianceFlags.voluntaryWithdrawal) status = "withdrawn";
  rsh.getRange(row, h["status"]).setValue(status);
  return { success: true };
}

// ============ ⑥ 事後アンケートの保存 ============
// demographics: { age, gender, nativeLanguage } は分析用に専用列へ、それ以外はpost_jsonへまとめて格納する。
function savePostQuestionnaire(sessionId, payload, demographics) {
  const rsh = getSheet_(SHEET.RESPONSES);
  const h = headerIndex_(rsh);
  const row = findRowByValue_(rsh, h["session_id"], sessionId);
  if (row === -1) return { success: false, reason: "セッションが見つかりません。" };

  demographics = demographics || {};
  rsh.getRange(row, h["age"]).setValue(demographics.age || "");
  rsh.getRange(row, h["gender"]).setValue(demographics.gender || "");
  rsh.getRange(row, h["native_language"]).setValue(demographics.nativeLanguage || "");
  rsh.getRange(row, h["post_json"]).setValue(JSON.stringify(payload || {}));
  rsh.getRange(row, h["post_saved_at"]).setValue(nowStr_());
  rsh.getRange(row, h["status"]).setValue("completed");

  return { success: true };
}
