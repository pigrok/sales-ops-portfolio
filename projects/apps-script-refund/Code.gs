const KST_TZ = "Asia/Seoul";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("환급 지표 계산기")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function ocrExtract(payload) {
  if (!payload) {
    throw new Error("이미지 데이터가 비어 있습니다.");
  }

  // 무료 Drive OCR을 유지하면서 카드별로 잘라 보낸 이미지를 각각 분석한다.
  // 각 이미지에는 라벨과 금액이 하나만 있으므로 전체 OCR의 순서 뒤섞임이 없다.
  if (Array.isArray(payload.cards) && payload.cards.length) {
    const fields = {};
    const texts = [];

    payload.cards.forEach((card, index) => {
      if (!card || !card.key || !card.base64) return;
      const blob = createImageBlob_(
        card.base64,
        card.mimeType || "image/jpeg",
        "card_" + card.key + "_" + index + ".jpg"
      );
      const text = driveOcr_(blob);
      const value = extractCardAmount_(text);
      texts.push("[" + card.key + "]\n" + text);
      if (Number.isFinite(value)) fields[card.key] = value;
    });

    if (Number.isFinite(fields.claimRefundIncluding)) {
      fields.claimRefundAssigned = fields.claimRefundIncluding;
    }
    return { ok: true, text: texts.join("\n\n"), fields: fields };
  }

  if (!payload.base64) {
    throw new Error("이미지 데이터가 비어 있습니다.");
  }
  const blob = createImageBlob_(
    payload.base64,
    payload.mimeType || "image/png",
    payload.fileName || ("upload_" + Date.now() + ".png")
  );

  const text = driveOcr_(blob);
  const fields = parseFields_(text);
  return { ok: true, text: text, fields: fields };
}

function createImageBlob_(base64, mimeType, fileName) {
  const bytes = Utilities.base64Decode(base64);
  return Utilities.newBlob(
    bytes,
    mimeType || "image/png",
    fileName || ("upload_" + Date.now() + ".png")
  );
}

function driveOcr_(blob) {
  const created = Drive.Files.create(
    {
      name: "ocr_tmp_" + Date.now(),
      mimeType: "application/vnd.google-apps.document"
    },
    blob,
    { ocrLanguage: "ko", fields: "id" }
  );

  try {
    const doc = DocumentApp.openById(created.id);
    return doc.getBody().getText() || "";
  } finally {
    try {
      Drive.Files.remove(created.id);
    } catch (e) {
      // 임시 파일 삭제 실패는 무시
    }
  }
}

/**
 * Drive OCR은 카드 좌표를 보존하지 않으므로 라벨 매칭을 사용하지 않는다.
 * 현재 대시보드에서 콤마가 포함된 KPI 금액 6개의 등장 순서로 고정 매핑한다.
 *
 * 순서(좌→우, 위→아래):
 *   [0] 전체 조회 환급액 (월)
 *   [1] 거래 전환된 총 조회환급액 (하반기)
 *   [2] 신청전환 성공 환급액
 *   [3] 신고완료 환급액
 *   [4] 취소 요청 금액
 *   [5] 취소방어 성공 환급액
 */
function parseFields_(rawText) {
  const amounts = extractDashboardAmounts_(rawText);
  const n = amounts.length;
  if (n < 6) return {};

  // OCR 원문에 앞쪽 숫자가 추가되더라도 마지막 KPI 6개를 기준으로 한다.
  const values = amounts.slice(n - 6);
  return {
    totalViewAmount: values[0],
    convertedViewRefund: values[1],
    successApplyViewRefund: values[2],
    claimRefundIncluding: values[3],
    claimRefundAssigned: values[3],
    cancelRequestAmount: values[4],
    successCancelDefenseViewRefund: values[5]
  };
}

/**
 * 날짜·건수·차트의 M 단위 값을 제외하기 위해 콤마가 있는 원 단위 금액만 추출한다.
 */
function extractDashboardAmounts_(rawText) {
  const text = String(rawText || "");
  // OCR이 마지막 세 자리를 분리하는 경우도 처리:
  //   "47,339,236 569" / "47,339,236, 569" → 47,339,236,569
  const re = /([0-9]{1,3}(?:,[0-9]{3})+)(?:[,\s]+([0-9]{3})(?!\d))?/g;
  const out = [];
  var m;
  while ((m = re.exec(text)) !== null) {
    const normalized = m[1].replace(/,/g, "") + (m[2] || "");
    const value = Number(normalized);
    if (isFinite(value) && value >= 1000000) out.push(value);
  }
  return out;
}

/** 카드 하나에서 가장 큰 원 단위 금액을 반환한다. */
function extractCardAmount_(rawText) {
  const amounts = extractDashboardAmounts_(rawText);
  if (!amounts.length) return null;
  return Math.max.apply(null, amounts);
}

const DAILY_HEADER = [
  "date",
  "totalViewAmount",
  "successApplyViewRefund",
  "successCancelDefenseViewRefund",
  "cancelRequestAmount",
  "convertedViewRefund",
  "claimRefundAssigned",
  "claimRefundIncluding",
  "conversionRate",
  "defenseRate",
  "claimRate",
  "totalConverted",
  "updatedAt",
  "userEmail"
];

const DATA_SPREADSHEET_ID = "1cybv-ZWAdBdyAocz8VEZkWryPrQ1Nz-RKrw6BDW8X90";

function getDataSpreadsheet_() {
  return SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
}

function getDailySheet_() {
  const ss = getDataSpreadsheet_();
  // DAILY를 단일 기준 시트로 사용한다.
  let sh = ss.getSheetByName("DAILY");
  if (!sh) {
    sh = ss.getSheetById(0) || ss.getSheets()[0];
    if (!sh) throw new Error("데이터를 저장할 시트를 찾을 수 없습니다.");
    sh.setName("DAILY");
  }

  migrateDailySheets_(ss, sh);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, DAILY_HEADER.length).setValues([DAILY_HEADER]);
  }
  sh.getRange("A:A").setNumberFormat("@");
  sh.getRange("M:M").setNumberFormat("@"); // updatedAt
  return sh;
}

/**
 * 예전 DAILY(9/12열)와 DAILY_V2(14열)를 현재 14열 구조로 통합한다.
 * 같은 날짜가 양쪽에 있으면 최신 구조인 DAILY_V2 데이터가 우선한다.
 * 성공적으로 통합한 뒤 DAILY_V2 탭은 삭제한다.
 */
function migrateDailySheets_(ss, daily) {
  const v2 = ss.getSheetByName("DAILY_V2");
  const dailyValues = daily.getLastRow() > 0
    ? daily.getDataRange().getValues()
    : [];
  const header = dailyValues.length ? dailyValues[0].map(String) : [];
  const headerIsCurrent = DAILY_HEADER.every((name, i) => header[i] === name);
  const hasLegacyRows = dailyValues.slice(1).some((row) => !isCurrentDailyRow_(row));

  if (!v2 && headerIsCurrent && !hasLegacyRows) return;

  const byDate = {};

  // 기존 DAILY를 먼저 변환한다.
  dailyValues.slice(1).forEach((row) => {
    const normalized = normalizeDailyRow_(row);
    if (normalized) byDate[normalized[0]] = normalized;
  });

  // DAILY_V2가 있으면 같은 날짜를 최신 데이터로 덮어쓴다.
  if (v2 && v2 !== daily && v2.getLastRow() > 1) {
    const v2Values = v2.getDataRange().getValues();
    v2Values.slice(1).forEach((row) => {
      const normalized = normalizeCurrentDailyRow_(row);
      if (normalized) byDate[normalized[0]] = normalized;
    });
  }

  const rows = Object.keys(byDate)
    .sort()
    .map((date) => byDate[date]);

  daily.clearContents();
  daily.getRange(1, 1, 1, DAILY_HEADER.length).setValues([DAILY_HEADER]);
  if (rows.length) {
    daily.getRange(2, 1, rows.length, DAILY_HEADER.length).setValues(rows);
  }
  SpreadsheetApp.flush();

  if (v2 && v2 !== daily) {
    ss.deleteSheet(v2);
  }
}

function isCurrentDailyRow_(row) {
  if (!normalizeDate_(row[0])) return true;
  return isTimestamp_(row[12]);
}

function isTimestamp_(value) {
  if (value instanceof Date) return true;
  return /^\d{4}-\d{2}-\d{2}[ T]/.test(String(value || ""));
}

function numberOrZero_(value) {
  const number = Number(value);
  return isFinite(number) ? number : 0;
}

function normalizeDailyRow_(row) {
  const date = normalizeDate_(row[0]);
  if (!date) return null;

  // 현재 14열 구조
  if (isCurrentDailyRow_(row)) return normalizeCurrentDailyRow_(row);

  // 구 12열 구조:
  // date, claim, converted, apply, cancelDefense, claimRate, total,
  // claimDefenseRequest, claimDefenseSuccess, includingRate, updatedAt, email
  if (isTimestamp_(row[10])) {
    const claim = numberOrZero_(row[1]);
    const converted = numberOrZero_(row[2]);
    const apply = numberOrZero_(row[3]);
    const cancelDefense = numberOrZero_(row[4]);
    const claimDefenseSuccess = numberOrZero_(row[8]);
    const claimIncluding = claim + claimDefenseSuccess;
    const claimRate = converted
      ? claimIncluding / converted
      : numberOrZero_(row[9] || row[5]);
    return [
      date, 0, apply, cancelDefense, 0, converted,
      claim, claimIncluding, 0, 0, claimRate,
      apply + cancelDefense, String(row[10] || ""), String(row[11] || "")
    ];
  }

  // 최초 9열 구조:
  // date, claim, converted, apply, cancelDefense, claimRate, total, updatedAt, email
  if (isTimestamp_(row[7])) {
    const claim = numberOrZero_(row[1]);
    const converted = numberOrZero_(row[2]);
    const apply = numberOrZero_(row[3]);
    const cancelDefense = numberOrZero_(row[4]);
    const claimRate = converted ? claim / converted : numberOrZero_(row[5]);
    return [
      date, 0, apply, cancelDefense, 0, converted,
      claim, claim, 0, 0, claimRate,
      apply + cancelDefense, String(row[7] || ""), String(row[8] || "")
    ];
  }

  // 구조를 판별할 수 없는 행은 수치 왜곡을 막기 위해 현재 위치 기준으로 정규화한다.
  return normalizeCurrentDailyRow_(row);
}

function normalizeCurrentDailyRow_(row) {
  const date = normalizeDate_(row[0]);
  if (!date) return null;
  const totalView = numberOrZero_(row[1]);
  const apply = numberOrZero_(row[2]);
  const cancelDefense = numberOrZero_(row[3]);
  const cancelRequest = numberOrZero_(row[4]);
  const converted = numberOrZero_(row[5]);
  const claimAssigned = numberOrZero_(row[6]);
  const claimIncluding = numberOrZero_(row[7]) || claimAssigned;

  return [
    date,
    totalView,
    apply,
    cancelDefense,
    cancelRequest,
    converted,
    claimAssigned,
    claimIncluding,
    totalView ? apply / totalView : 0,
    cancelRequest ? cancelDefense / cancelRequest : 0,
    converted ? claimIncluding / converted : 0,
    apply + cancelDefense,
    formatTimestamp_(row[12]),
    String(row[13] || "")
  ];
}

function formatTimestamp_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, KST_TZ, "yyyy-MM-dd HH:mm:ss");
  }
  return String(value || "");
}

function ensureDailyHeader_(sh) {
  const header = sh.getRange(1, 1, 1, DAILY_HEADER.length).getValues()[0].map(String);
  if (!DAILY_HEADER.every((name, i) => header[i] === name)) {
    sh.getRange(1, 1, 1, DAILY_HEADER.length).setValues([DAILY_HEADER]);
  }
}

function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, KST_TZ, "yyyy-MM-dd");
  return String(v == null ? "" : v).trim();
}

/** 이메일 권한이 없어도 저장 자체는 실패하지 않도록 한다. */
function getUserEmail_() {
  try {
    const active = Session.getActiveUser().getEmail();
    if (active) return active;
  } catch (e) {
    // 권한 미승인 시 무시하고 다음 방법을 시도한다.
  }
  try {
    return Session.getEffectiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function saveDaily(payload) {
  if (!payload) throw new Error("저장할 데이터가 없습니다.");

  const date = String((payload && payload.date) || "").trim();
  if (!date) throw new Error("기준일자가 비어 있습니다.");

  const totalView = Number(payload.totalViewAmount) || 0;
  const apply = Number(payload.successApplyViewRefund) || 0;
  const cancelDef = Number(payload.successCancelDefenseViewRefund) || 0;
  const cancelReq = Number(payload.cancelRequestAmount) || 0;
  const converted = Number(payload.convertedViewRefund) || 0;
  const claimAssigned = Number(payload.claimRefundAssigned) || 0;
  const claimIncluding = Number(payload.claimRefundIncluding) || 0;

  const conversionRate = totalView ? apply / totalView : 0;
  const defenseRate = cancelReq ? cancelDef / cancelReq : 0;
  const claimRate = converted ? claimIncluding / converted : 0;
  const totalConverted = apply + cancelDef;
  const now = Utilities.formatDate(new Date(), KST_TZ, "yyyy-MM-dd HH:mm:ss");
  const email = getUserEmail_();

  const row = [
    date,
    totalView,
    apply,
    cancelDef,
    cancelReq,
    converted,
    claimAssigned,
    claimIncluding,
    conversionRate,
    defenseRate,
    claimRate,
    totalConverted,
    now,
    email
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getDailySheet_();
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (normalizeDate_(values[i][0]) === date) {
        sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
        SpreadsheetApp.flush();
        return {
          ok: true,
          updated: true,
          date: date,
          row: i + 1,
          sheetName: sh.getName()
        };
      }
    }

    const targetRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    return {
      ok: true,
      updated: false,
      date: date,
      row: targetRow,
      sheetName: sh.getName()
    };
  } finally {
    lock.releaseLock();
  }
}

function listDaily(yearMonth) {
  const sh = getDailySheet_();
  const values = sh.getDataRange().getValues();
  const ym = String(yearMonth || "").trim();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const date = normalizeDate_(values[i][0]);
    if (!date) continue;
    if (ym && date.slice(0, 7) !== ym) continue;
    out.push({
      date: date,
      totalViewAmount: Number(values[i][1]) || 0,
      successApplyViewRefund: Number(values[i][2]) || 0,
      successCancelDefenseViewRefund: Number(values[i][3]) || 0,
      cancelRequestAmount: Number(values[i][4]) || 0,
      convertedViewRefund: Number(values[i][5]) || 0,
      claimRefundAssigned: Number(values[i][6]) || 0,
      claimRefundIncluding: Number(values[i][7]) || 0,
      conversionRate: Number(values[i][8]) || 0,
      defenseRate: Number(values[i][9]) || 0,
      claimRate: Number(values[i][10]) || 0,
      totalConverted: Number(values[i][11]) || 0,
      updatedAt: String(values[i][12] || ""),
      userEmail: String(values[i][13] || "")
    });
  }
  out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return out;
}

function deleteDaily(date) {
  const target = String(date || "").trim();
  if (!target) throw new Error("삭제할 날짜가 없습니다.");
  const sh = getDailySheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalizeDate_(values[i][0]) === target) {
      sh.deleteRow(i + 1);
      return { ok: true, date: target };
    }
  }
  return { ok: false, date: target };
}
