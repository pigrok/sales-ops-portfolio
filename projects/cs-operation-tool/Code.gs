const KST_TZ = "Asia/Seoul";


function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("CS Operation Tool")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitData() {
  const email = Session.getActiveUser().getEmail() || "";
  const user = requireAllowedUser_(email);
  const targets = getPodTargets_();
  const assignees = getAssignees_();
  const activityTypes = getActivityTypes_();
  return { 
    ok: true,
    email: email,
    userName: user.name || "",
    defaultAssigneeUserId: user.assigneeUserId || "",
    assignees: assignees,
    targets: targets,
    activityTypes: activityTypes,
  };
}

function transcribeCall(req) {
  const email = Session.getActiveUser().getEmail() || "";
  const user = requireAllowedUser_(email);
  validateReq_(req, ["fileName", "mimeType", "base64Audio"]);

  const decoded = Utilities.base64Decode(req.base64Audio || "");
  if (!decoded || decoded.length === 0) {
    throw new Error("파일을 읽을 수 없습니다. 파일을 다시 선택한 후 업로드해 주세요.");
  }
  const ext = String(req.fileName || "").split(".").pop() || "m4a";
  const audioBlob = Utilities.newBlob(decoded, req.mimeType, "recording." + ext);

  return transcribeAudioBlob_(audioBlob, user);
}

function transcribeCallFromDrive(req) {
  const email = Session.getActiveUser().getEmail() || "";
  const user = requireAllowedUser_(email);
  validateReq_(req, ["fileId"]);

  const fileId = String(req.fileId || "").trim();
  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    throw new Error("녹음 파일에 접근할 수 없습니다. 본인 Google Drive 권한을 확인해 주세요.");
  }
  if (!isAudioFile_(file)) {
    throw new Error("오디오 파일이 아닙니다.");
  }

  const ext = file.getName().split(".").pop() || "m4a";
  const blob = file.getBlob().setName("recording." + ext);
  const result = transcribeAudioBlob_(blob, user);
  result.fileId = fileId;
  result.fileName = file.getName();
  return result;
}

function listDriveRecordings(req) {
  const email = Session.getActiveUser().getEmail() || "";
  const user = requireAllowedUser_(email);
  const limit = Math.min(Math.max(Number((req || {}).limit || 15), 1), 30);

  const folderId = resolveRecordingsFolderId_(user);
  let files = [];
  let source = "recent";

  if (folderId) {
    try {
      files = listAudioFilesFromFolder_(DriveApp.getFolderById(folderId), limit, 300);
      source = "folder";
    } catch (e) {
      files = [];
    }
  }
  if (!files.length) {
    files = listRecentAudioInDrive_(limit);
    source = "recent";
  }

  files.sort(function (a, b) {
    return Number(b.modifiedMs || 0) - Number(a.modifiedMs || 0);
  });

  const driveCtx = getDriveExecutionContext_();

  return {
    ok: true,
    files: files.slice(0, limit),
    folderId: folderId || "",
    source: source,
    driveAccountEmail: driveCtx.effectiveEmail,
    loginEmail: driveCtx.activeEmail,
    driveMismatch: driveCtx.mismatch,
  };
}

function transcribeAudioBlob_(audioBlob, user) {
  const stt = transcribeAudioWithGroq_(audioBlob);
  const transcript = String(stt.text || "");
  var dialogue = "";
  var formatError = null;
  try {
    dialogue = formatDialogueWithGroq_(transcript);
  } catch (e) {
    formatError = String(e.message || e);
    dialogue = transcript;
  }
  return {
    ok: true,
    transcript: transcript,
    dialogue: dialogue,
    durationSec: Number(stt.duration || 0),
    userName: user.name || "",
    formatError: formatError,
  };
}

function createCallActivity(req) {
  const email = Session.getActiveUser().getEmail() || "";
  const user = requireAllowedUser_(email);

  validateReq_(req, [
    "recordRef",
    "callDatetimeKst",
    "subject",
    "assigneeUserId",
    "activityTypeKey",
    "targetSystem",
    "corpPhone",
    "dialogue",
  ]);

  if (!isAssigneeAllowedForTarget_(String(req.targetSystem || ""), String(req.assigneeUserId || ""))) {
    throw new Error("선택한 대상 시스템에서 사용할 수 없는 담당자 ID입니다. ASSIGNEES 시트를 확인해 주세요.");
  }

  const token = resolvePipedriveToken_(String(req.targetSystem || ""));
  const baseUrl = getProp_("PIPEDRIVE_BASE_URL", "https://api.pipedrive.com/v1");
  const normalizedCorpPhone = normalizeCorpPhone_(String(req.corpPhone || ""));
  const parsedRecord = parseRecordRef_(String(req.recordRef || ""));
  const recordType = parsedRecord.recordType;
  const recordId = parsedRecord.recordId;

  const durationSec = Number(req.durationSec || 0);
  const dialogue = String(req.dialogue || "");

  const due = toUtcDue_(req.callDatetimeKst);
  const note = buildActivityNote_({
    createdAt: Utilities.formatDate(new Date(), KST_TZ, "yyyy-MM-dd HH:mm:ss"),
    operatorName: user.name || "",
    operatorPhone: user.phone || "",
    callDatetimeKst: req.callDatetimeKst,
    durationSec: durationSec,
    specialNote: req.specialNote || "",
    dialogue: dialogue,
  });

  const body = {
    subject: String(req.subject || "통화 기록"),
    type: String(req.activityTypeKey || "call"),
    done: 1,
    due_date: due.date,
    due_time: due.time,
    note: note,
  };
  if (recordType === "LEAD") {
    body.lead_id = recordId;
  } else {
    // 기본: DEAL
    if (!/^\d+$/.test(recordId)) {
      throw new Error("거래(DEAL) 유형은 숫자 ID만 입력 가능합니다.");
    }
    body.deal_id = Number(recordId);
  }
  if (String(req.assigneeUserId || "").match(/^\d+$/)) {
    body.user_id = Number(req.assigneeUserId);
  }

  const created = pdPost_(baseUrl, token, "/activities", body);
  const activityId = Number((created.data || {}).id || 0);
  if (!activityId) {
    throw new Error("활동 생성 실패: activity id 없음");
  }

  // 취소방어: 담당자 커스텀 필드 자동 채우기 (DEAL만)
  var cancelDefenseResult = { skipped: true };
  if (String(req.activityTypeKey || "") === "code39881401" && recordType === "DEAL") {
    try {
      cancelDefenseResult = updateCancelDefenseAssignee_(baseUrl, token, recordId, String(req.assigneeUserId || ""));
    } catch (e) {
      cancelDefenseResult = { skipped: true, error: e.message };
    }
  }

  const uploadResult = { ok: true, skipped: true, reason: "녹취 첨부 비활성화" };

  appendCallLog_({
    createdAt: Utilities.formatDate(new Date(), KST_TZ, "yyyy-MM-dd HH:mm:ss"),
    operatorName: user.name || "",
    operatorPhone: user.phone || "",
    callDatetimeKst: String(req.callDatetimeKst || ""),
    callDurationSec: durationSec,
    selectedAssignee: String(req.assigneeUserId || ""),
    mySelection: String(req.activityTypeName || req.activityTypeKey || ""),
    corpPhone: normalizedCorpPhone,
    targetSystem: String(req.targetSystem || ""),
    dealId: `${recordType}:${recordId}`,
    activityId: activityId,
    userEmail: email,
  });

  return {
    ok: true,
    activityId: activityId,
    uploadResult: uploadResult,
    dialogue: dialogue,
    activityTypeKey: String(req.activityTypeKey || ""),
    recordType: recordType,
    recordId: recordId,
    cancelDefenseResult: cancelDefenseResult,
  };
}

function parseRecordRef_(raw) {
  const s = String(raw || "").trim();
  if (!s) throw new Error("리드/거래 URL 또는 ID를 입력해 주세요.");

  // DEAL=12345 / LEAD=uuid 형식 지원
  const pref = s.match(/^(DEAL|LEAD)\s*[:=]\s*(.+)$/i);
  if (pref) {
    const t = pref[1].toUpperCase();
    const id = String(pref[2] || "").trim();
    if (!id) throw new Error("리드/거래 ID가 비어 있습니다.");
    if (t === "DEAL" && !/^\d+$/.test(id)) {
      throw new Error("거래(DEAL) ID는 숫자만 가능합니다.");
    }
    return { recordType: t, recordId: id };
  }

  // deal URL
  let m = s.match(/\/deal\/(\d+)(?:[/?#]|$)/i);
  if (m) {
    return { recordType: "DEAL", recordId: String(m[1]) };
  }

  // lead URL
  m = s.match(/\/leads\/inbox\/([^/?#]+)(?:[/?#]|$)/i);
  if (m) {
    return { recordType: "LEAD", recordId: String(m[1]) };
  }

  // 숫자만이면 DEAL
  if (/^\d+$/.test(s)) {
    return { recordType: "DEAL", recordId: s };
  }

  // UUID면 LEAD
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s)) {
    return { recordType: "LEAD", recordId: s };
  }

  throw new Error("리드/거래 URL 또는 ID 형식이 올바르지 않습니다.");
}

function getActivityTypes_() {
  // 요구사항: 2개만 고정
  return [
    { key: "code371194033", name: "Refund Recontact Campaign" },
    { key: "code39881401", name: "취소방어" },
  ];
}

function requireAllowedUser_(email) {
  if (!email) throw new Error("로그인 이메일을 확인할 수 없습니다.");
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("USERS");
  if (!sh) throw new Error("USERS 시트를 만들어 주세요.");
  const values = sh.getDataRange().getValues();
  // header: email, name, phone, enabled, assignee_user_id(optional), recordings_folder_id(optional)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = String(row[0] || "").trim().toLowerCase();
    const enabled = String(row[3] || "Y").trim().toUpperCase();
    if (rowEmail === email.toLowerCase() && enabled === "Y") {
      return {
        email: rowEmail,
        name: String(row[1] || ""),
        phone: String(row[2] || ""),
        assigneeUserId: String(row[4] || "").trim(),
        recordingsFolderId: String(row[5] || "").trim(),
      };
    }
  }
  throw new Error("사용 권한이 없습니다. 관리자에게 USERS 시트 등록 요청하세요.");
}

function getAssignees_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("ASSIGNEES");
  if (!sh) return [];
  const values = sh.getDataRange().getDisplayValues();
  const out = [];
  // header: user_id | name | target_key(optional) | corp_phone(optional)
  for (let i = 1; i < values.length; i++) {
    const userId = String(values[i][0] || "").trim();
    const name = String(values[i][1] || "").trim();
    const targetKey = String(values[i][2] || "").trim().toUpperCase();
    const corpPhone = normalizeCorpPhone_(String(values[i][3] || "").trim());
    if (userId && name) out.push({ userId: userId, name: name, targetKey: targetKey, corpPhone: corpPhone });
  }
  return out;
}

function normalizeCorpPhone_(raw) {
  let digits = String(raw || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.startsWith("82")) digits = "0" + digits.slice(2);
  // 1058058843 -> 010-XXXX-XXXX (모바일 앞 0 유실 보정)
  if (digits.length === 10 && digits.startsWith("10")) digits = "0" + digits;
  return digits;
}

function isAssigneeAllowedForTarget_(targetSystem, assigneeUserId) {
  const target = String(targetSystem || "").trim().toUpperCase();
  const uid = String(assigneeUserId || "").trim();
  if (!target || !uid) return false;
  const assignees = getAssignees_();
  const matched = assignees.filter(function (a) {
    return String(a.userId || "").trim() === uid;
  });
  if (matched.length === 0) return false;
  // targetKey가 비어있는 담당자는 공통(모든 타겟 허용)
  return matched.some(function (a) {
    const tk = String(a.targetKey || "").trim().toUpperCase();
    return tk === "" || tk === target;
  });
}

function getTargets_() {
  return getPodTargets_();
}

function getPodTargets_() {
  // 1순위: TARGETS 시트(key|label)
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("TARGETS");
  if (sh) {
    const values = sh.getDataRange().getValues();
    const fromSheet = [];
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || "").trim().toUpperCase();
      const label = String(values[i][1] || "").trim();
      if (key && label) fromSheet.push({ key: key, label: label });
    }
    if (fromSheet.length > 0) return fromSheet;
  }

  // 2순위: Script Properties POD_TARGETS_JSON
  const defaults = [
    { key: "CARE", label: "케어" },
    { key: "REFUND", label: "환급" },
  ];

  const json = getProp_("POD_TARGETS_JSON", "");
  if (!json) return defaults;
  try {
    const arr = JSON.parse(json);
    if (!Array.isArray(arr) || arr.length === 0) return defaults;
    const out = [];
    arr.forEach(function (x) {
      const key = String((x || {}).key || "").trim().toUpperCase();
      const label = String((x || {}).label || "").trim();
      if (key && label) out.push({ key: key, label: label });
    });
    return out.length ? out : defaults;
  } catch (e) {
    return defaults;
  }
}

function resolveRecordingsFolderId_(user) {
  const fromUser = String((user || {}).recordingsFolderId || "").trim();
  if (fromUser) return fromUser;

  // 담당자별 폴더 미지정 시: 로그인 사용자 Drive에서만 폴더명 탐색 (공통 폴더 ID 사용 안 함)
  const names = [
    getProp_("RECORDINGS_FOLDER_NAME", "통화녹음"),
    "통화 녹음",
    "Call Recordings",
    "녹음",
    "recordings",
  ];
  const seen = {};
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i] || "").trim();
    if (!name || seen[name]) continue;
    seen[name] = true;
    const found = findFolderByName_(name);
    if (found) return found;
  }
  return "";
}

function getDriveExecutionContext_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || "").trim();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || "").trim();
  const mismatch =
    !!activeEmail &&
    !!effectiveEmail &&
    activeEmail.toLowerCase() !== effectiveEmail.toLowerCase();
  return {
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    mismatch: mismatch,
  };
}

function findFolderByName_(name) {
  const it = DriveApp.getFoldersByName(String(name || "").trim());
  if (it.hasNext()) return it.next().getId();
  return "";
}

function isAudioFile_(file) {
  const mime = String(file.getMimeType() || "").toLowerCase();
  const name = String(file.getName() || "").toLowerCase();
  if (mime.indexOf("audio/") === 0) return true;
  return /\.(m4a|mp3|wav|mp4|mpeg|mpga|webm|flac|ogg)$/i.test(name);
}

function toRecordingMeta_(file) {
  const updated = file.getLastUpdated();
  return {
    id: file.getId(),
    name: file.getName(),
    modifiedAt: Utilities.formatDate(updated, KST_TZ, "yyyy-MM-dd HH:mm:ss"),
    modifiedMs: updated.getTime(),
    size: file.getSize(),
    mimeType: file.getMimeType() || "",
  };
}

function listAudioFilesFromFolder_(folder, limit, maxScan) {
  const out = [];
  const it = folder.getFiles();
  let scanned = 0;
  while (it.hasNext() && out.length < limit && scanned < maxScan) {
    scanned++;
    const f = it.next();
    if (!isAudioFile_(f)) continue;
    out.push(toRecordingMeta_(f));
  }
  return out;
}

function listRecentAudioInDrive_(limit) {
  const out = [];
  const since = new Date();
  since.setDate(since.getDate() - 14);
  const sinceText = Utilities.formatDate(since, "GMT", "yyyy-MM-dd'T'HH:mm:ss");
  const it = DriveApp.searchFiles('modifiedDate > "' + sinceText + '" and trashed = false');
  let scanned = 0;
  while (it.hasNext() && out.length < limit && scanned < 500) {
    scanned++;
    const f = it.next();
    if (!isAudioFile_(f)) continue;
    out.push(toRecordingMeta_(f));
  }
  return out;
}

function resolvePipedriveToken_(targetSystem) {
  const mapJson = getProp_("TARGET_PIPEDRIVE_TOKENS_JSON", "{}");
  let mapObj = {};
  try {
    mapObj = JSON.parse(mapJson);
  } catch (e) {
    mapObj = {};
  }
  const key = String(targetSystem || "").trim().toUpperCase();
  const token = String(mapObj[key] || mapObj[String(targetSystem || "").trim()] || "").trim();
  if (!token) throw new Error("TARGET_PIPEDRIVE_TOKENS_JSON에 대상 토큰이 없습니다. target=" + key);
  return token;
}

function buildActivityNote_(data) {
  const createdAt = String(data.createdAt || "").trim();
  const operatorName = String(data.operatorName || "").trim();
  const operatorPhone = String(data.operatorPhone || "").trim();
  const special = String(data.specialNote || "").trim();
  const dialogue = String(data.dialogue || "").trim();

  const dialogueHtml = dialogue
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");

  const blocks = [
    `${createdAt} / ${operatorName} / ${operatorPhone}`,
    special ? `[특이사항] ${special}` : "[특이사항 없음]",
  ];
  if (dialogue) {
    blocks.push(`[대화 정리]<br>${dialogueHtml}`);
  } else {
    blocks.push("[대화 정리 없음]");
  }
  return blocks.join("<br><br>");
}

function toUtcDue_(kstDateTimeText) {
  // input: yyyy-MM-ddTHH:mm
  const t = String(kstDateTimeText || "").trim();
  if (!t) throw new Error("통화 시각이 비어 있습니다.");
  const dt = new Date(t + ":00+09:00");
  return {
    date: Utilities.formatDate(dt, "UTC", "yyyy-MM-dd"),
    time: Utilities.formatDate(dt, "UTC", "HH:mm"),
  };
}

function secToMMSS_(sec) {
  const s = Math.max(0, Math.round(Number(sec || 0)));
  const m = Math.floor(s / 60);
  const r = s % 60;
  return (m < 10 ? "0" + m : String(m)) + ":" + (r < 10 ? "0" + r : String(r));
}

function appendCallLog_(row) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("CALL_LOG");
  if (!sh) throw new Error("CALL_LOG 시트를 만들어 주세요.");
  sh.appendRow([
    row.createdAt,
    row.operatorName,
    row.operatorPhone,
    row.callDatetimeKst,
    row.callDurationSec,
    row.selectedAssignee,
    row.mySelection,
    row.corpPhone,
    row.targetSystem,
    row.dealId,
    row.activityId,
    row.userEmail,
  ]);
}

function transcribeAudioWithGroq_(audioBlob) {
  const apiKey = getProp_("GROQ_API_KEY", "");
  const url = "https://api.groq.com/openai/v1/audio/transcriptions";
  const fields = {
    model: getProp_("OPENAI_WHISPER_MODEL", "whisper-large-v3-turbo"),
    language: "ko",
    response_format: "verbose_json",
    "timestamp_granularities[]": "segment",
  };
  const res = multipartFetch_(url, apiKey, fields, audioBlob);
  return {
    text: String(res.text || ""),
    duration: Number(res.duration || 0),
  };
}

function dialogueSystemPrompt_() {
  return (
    "너는 한국어 영업·CS 통화 녹취의 화자 분리 전문가다.\n" +
    "Tax SaaS(세금 환급 서비스) 매니저가 사장님(고객)에게 아웃바운드 전화를 건 상황이다.\n" +
    "입력으로 화자 구분 없이 이어진 한국어 통화 전사문이 주어진다.\n\n" +
    "출력 형식:\n" +
    "- 모든 줄은 반드시 '매니저: ' 또는 '고객: ' 으로 시작한다.\n" +
    "- 라벨 뒤에 해당 화자의 발화를 원문 그대로 적는다.\n" +
    "- 화자가 바뀔 때마다 새 줄로 나눈다.\n\n" +
    "화자 판별 기준 (우선순위 순):\n" +
    "1. '저희(회사·서비스)', '해드리다', '안내드리다', '진행해드리다', '도와드리다' 포함 → 매니저\n" +
    "2. 상대방을 '사장님'으로 호칭 → 매니저\n" +
    "3. 환급·검토·신청·서비스 내용을 설명하거나 권유 → 매니저\n" +
    "4. 짧은 수신 반응('네', '아 네', '그렇군요', '음', '아 그래요') → 앞 화자의 반대편\n" +
    "5. 의문·거절·개인 상황 언급('돈 넣잖아요', '고민해볼게요', '세무사님 통해서 했어요') → 고객\n" +
    "6. 전체 흐름에서 역할을 일관되게 유지 (매니저는 전화를 건 쪽, 고객은 받는 쪽)\n\n" +
    "규칙:\n" +
    "- 원문 단어를 절대 바꾸지 않는다. 라벨만 붙인다.\n" +
    "- 요약·삭제·추가·수정 금지.\n" +
    "- 한 화자가 연속으로 여러 문장을 말할 수 있다.\n\n" +
    "출력 예시:\n" +
    "매니저: 안녕하세요 사장님, Tax Refund SaaS 담당자입니다.\n" +
    "고객: 네, 안녕하세요.\n" +
    "매니저: 환급 관련해서 연락드렸는데 잠깐 통화 가능하실까요?\n" +
    "고객: 네 가능합니다.\n" +
    "매니저: 사장님 같은 경우에 5년치 종합소득세 신고 내역에서 누락된 감면 항목이 있어서요.\n" +
    "고객: 아 그래요?\n" +
    "매니저: 네, 저희가 무료로 2차 검토를 해드리고 있어서 신청 한번 해보시는 게 어떨까 합니다.\n" +
    "고객: 지금 다른 데서 이미 진행하고 있는데요."
  );
}

function cleanDialogueOutput_(raw) {
  const cleaned = String(raw || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, "")
    .replace(/<think>[\s\S]*/i, "")
    .replace(/```[^\n]*\n([\s\S]*?)```/g, "$1")
    .trim();
  const m = cleaned.match(/((?:매니저|고객)\s*:[\s\S]+)/);
  return m ? m[1].trim() : cleaned;
}



function estimateMaxTokens_(transcript) {
  // 한국어 1자당 약 2 토큰, 화자 라벨 추가 30% 여유, 최소 800 / 최대 4096
  return Math.min(Math.max(Math.ceil(transcript.length * 2 * 1.3) + 200, 800), 4096);
}

function callCloudflare_(transcript) {
  const apiToken = getProp_("CLOUDFLARE_API_TOKEN", "");
  const accountId = getProp_("CLOUDFLARE_ACCOUNT_ID", "");
  if (!apiToken || !accountId) return null;
  const url = "https://api.cloudflare.com/client/v4/accounts/" + accountId + "/ai/v1/chat/completions";
  const models = [
    "@cf/meta/llama-3.3-70b-instruct-fp8-fast",
    "@cf/meta/llama-3.1-70b-instruct",
    "@cf/qwen/qwen1.5-14b-chat-awq",
    "@cf/meta/llama-3.1-8b-instruct",
  ];
  const maxTokens = estimateMaxTokens_(transcript);
  for (let i = 0; i < models.length; i++) {
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiToken },
      payload: JSON.stringify({
        model: models[i],
        temperature: 0,
        max_tokens: maxTokens,
        messages: [
          { role: "system", content: dialogueSystemPrompt_() },
          { role: "user", content: transcript },
        ],
      }),
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log("[Cloudflare/" + models[i] + "] " + code + ": " + resp.getContentText().slice(0, 200));
      continue;
    }
    try {
      const json = JSON.parse(resp.getContentText());
      const content = String((((json.choices || [])[0] || {}).message || {}).content || "").trim();
      if (!content) {
        Logger.log("[Cloudflare/" + models[i] + "] 빈 응답");
        continue;
      }
      if (/<think>/i.test(content) && !/<\/think>/i.test(content)) {
        Logger.log("[Cloudflare/" + models[i] + "] <think> 잘림");
        continue;
      }
      const cleaned = cleanDialogueOutput_(content);
      if (!/(?:매니저|고객)\s*:/.test(cleaned)) {
        Logger.log("[Cloudflare/" + models[i] + "] 화자 형식 없음. 앞 200자: " + cleaned.slice(0, 200));
        continue;
      }
      return content;
    } catch (e) {
      Logger.log("[Cloudflare/" + models[i] + "] 파싱 실패: " + e.message);
    }
  }
  return null;
}

function formatDialogueWithGroq_(transcript) {
  if (!transcript) return "";

  // 1순위: Cloudflare Workers AI
  try {
    const cfRaw = callCloudflare_(transcript);
    if (cfRaw) {
      const dialogue = cleanDialogueOutput_(cfRaw);
      if (dialogue && /(?:매니저|고객)\s*:/.test(dialogue)) return dialogue;
    }
  } catch (e) {}

  // 2순위: Groq 폴백
  const result = groqChatWithFallback_({
    temperature: 0.0,
    max_tokens: estimateMaxTokens_(transcript),
    messages: [
      { role: "system", content: dialogueSystemPrompt_() },
      { role: "user", content: transcript },
    ],
  });
  const dialogue = cleanDialogueOutput_(result);
  if (!dialogue) throw new Error("LLM이 빈 응답을 반환했습니다.");
  if (!/(?:매니저|고객)\s*:/.test(dialogue)) throw new Error("LLM 응답에 화자 구분 형식(매니저:/고객:)이 없습니다.");
  return dialogue;
}

function getChatModelCandidates_() {
  const primary = getProp_("OPENAI_CHAT_MODEL", "llama-3.3-70b-versatile");
  const fallbackRaw = getProp_("OPENAI_CHAT_MODEL_FALLBACKS", "");
  const list = [String(primary || "").trim()];

  if (fallbackRaw) {
    String(fallbackRaw).split(",").forEach(function (m) {
      const name = String(m || "").trim();
      if (name && list.indexOf(name) < 0) list.push(name);
    });
  } else {
    try {
      const apiKey = getProp_("GROQ_API_KEY", "");
      const resp = UrlFetchApp.fetch("https://api.groq.com/openai/v1/models", {
        headers: { Authorization: "Bearer " + apiKey },
        muteHttpExceptions: true,
      });
      if (resp.getResponseCode() === 200) {
        const models = JSON.parse(resp.getContentText()).data || [];
        models
          .filter(function (m) {
            const id = String(m.id || "");
            return !/whisper|tts|vision|guard|distil|compound|orpheus|canopy|speech|audio|allam|gpt-oss|8b-instant/i.test(id);
          })
          .sort(function (a, b) {
            return Number(b.context_window || 0) - Number(a.context_window || 0);
          })
          .forEach(function (m) {
            const name = String(m.id || "").trim();
            if (name && list.indexOf(name) < 0) list.push(name);
          });
      }
    } catch (e) {}
  }

  [
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "meta-llama/llama-4-maverick-17b-128e-instruct",
    "llama-3.1-70b-versatile",
    "llama3-70b-8192",
    "mixtral-8x7b-32768",
  ].forEach(function (m) {
    if (list.indexOf(m) < 0) list.push(m);
  });

  return list.filter(Boolean);
}

function groqChatWithFallback_(basePayload) {
  const apiKey = getProp_("GROQ_API_KEY", "");
  const url = "https://api.groq.com/openai/v1/chat/completions";
  const models = getChatModelCandidates_();
  const errLog = [];

  for (let i = 0; i < models.length; i++) {
    const payload = {};
    Object.keys(basePayload).forEach(function (k) {
      payload[k] = basePayload[k];
    });
    payload.model = models[i];
    if (/qwq|deepseek|[-_]r1[-_]/i.test(models[i])) {
      payload.reasoning_format = "hidden";
    }

    let attempt413 = 0;
    let resp, code, text;
    while (true) {
      resp = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + apiKey },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      code = resp.getResponseCode();
      text = resp.getContentText();
      if (code === 413 && attempt413 < 4 && payload.max_tokens > 512) {
        payload.max_tokens = Math.floor(payload.max_tokens / 2);
        attempt413++;
        continue;
      }
      break;
    }

    if (code >= 200 && code < 300) {
      const json = JSON.parse(text);
      const content = String((((json.choices || [])[0] || {}).message || {}).content || "").trim();
      if (content) {
        if (/<think>/i.test(content) && !/<\/think>/i.test(content)) {
          Logger.log("[Groq/" + models[i] + "] <think> 잘림");
          errLog.push("[" + models[i] + "] <think> 잘림");
          continue;
        }
        return content;
      }
      const finishReason = ((json.choices || [])[0] || {}).finish_reason || "unknown";
      Logger.log("[Groq/" + models[i] + "] 빈 응답 (finish_reason=" + finishReason + ")");
      errLog.push("[" + models[i] + "] 빈 응답 (finish_reason=" + finishReason + ")");
      continue;
    }

    Logger.log("[Groq/" + models[i] + "] HTTP" + code + ": " + text.slice(0, 200));
    errLog.push("[" + models[i] + "] HTTP" + code);
    const retryable = code === 429 || code === 404 || code === 413 ||
      /rate_limit|model_not_found|does not exist|decommissioned|deprecated|not supported|max_tokens|less than or equal|terms/i.test(text);
    if (!retryable) break;
  }

  throw new Error("Groq 화자분리 실패 (" + errLog.length + "개 모델): " + errLog.join(" / "));
}

function pdGet_(baseUrl, token, path) {
  const url = baseUrl.replace(/\/+$/, "") + path + "?api_token=" + encodeURIComponent(token);
  const resp = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("Pipedrive GET 실패(" + code + "): " + text);
  }
  return JSON.parse(text);
}

function pdPut_(baseUrl, token, path, body) {
  const url = baseUrl.replace(/\/+$/, "") + path + "?api_token=" + encodeURIComponent(token);
  const resp = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("Pipedrive PUT 실패(" + code + "): " + text);
  }
  return JSON.parse(text);
}

function getCancelDefenseFieldKey_(baseUrl, token) {
  const fromProp = getProp_("CANCEL_DEFENSE_PERSON_FIELD_KEY", "");
  if (fromProp) return fromProp;
  try {
    const resp = pdGet_(baseUrl, token, "/dealFields");
    const fields = resp.data || [];
    for (var i = 0; i < fields.length; i++) {
      if (String((fields[i] || {}).name || "").indexOf("취소방어 담당자") >= 0) {
        return String((fields[i] || {}).key || "");
      }
    }
  } catch (e) {}
  return "";
}

function updateCancelDefenseAssignee_(baseUrl, token, dealId, assigneeUserId) {
  const fieldKey = getCancelDefenseFieldKey_(baseUrl, token);
  if (!fieldKey) return { skipped: true, reason: "취소방어 담당자 필드 키를 찾지 못했습니다." };

  const deal = pdGet_(baseUrl, token, "/deals/" + dealId);
  const currentVal = (deal.data || {})[fieldKey];
  if (currentVal !== null && currentVal !== undefined && String(currentVal || "").trim() !== "") {
    return { skipped: true, reason: "이미 담당자 설정됨" };
  }

  const updateBody = {};
  updateBody[fieldKey] = Number(assigneeUserId);
  pdPut_(baseUrl, token, "/deals/" + dealId, updateBody);
  return { ok: true };
}

function pdPost_(baseUrl, token, path, body) {
  const url = baseUrl.replace(/\/+$/, "") + path + "?api_token=" + encodeURIComponent(token);
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("Pipedrive POST 실패(" + code + "): " + text);
  }
  return JSON.parse(text);
}

function uploadFileToActivity_(baseUrl, token, activityId, fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  return uploadBlobToActivity_(baseUrl, token, activityId, blob);
}

function uploadBlobToActivity_(baseUrl, token, activityId, blob) {
  const url = baseUrl.replace(/\/+$/, "") + "/files";
  const fields = { activity_id: String(activityId) };
  const res = multipartFetch_(
    url + "?api_token=" + encodeURIComponent(token),
    "",
    fields,
    blob,
    false
  );
  return { ok: true, result: res };
}

function saveTempAudio_(blob) {
  const folderId = getProp_("TEMP_AUDIO_FOLDER_ID", "");
  if (!folderId) throw new Error("Script Properties에 TEMP_AUDIO_FOLDER_ID 설정이 필요합니다.");
  const folder = DriveApp.getFolderById(folderId);
  const f = folder.createFile(blob);
  return f.getId();
}

function multipartFetch_(url, bearerToken, fields, fileBlob, withAuthHeader) {
  const boundary = "----csOpBoundary" + new Date().getTime();
  const payload = buildMultipartPayload_(boundary, fields, fileBlob);
  const headers = {};
  if (withAuthHeader !== false && bearerToken) {
    headers.Authorization = "Bearer " + bearerToken;
  }

  while (true) {
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      headers: headers,
      contentType: "multipart/form-data; boundary=" + boundary,
      payload: payload,
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code === 429) {
      const m = text.match(/try again in (\d+(?:\.\d+)?)s/i);
      const wait = m ? Math.ceil(Number(m[1])) : 30;
      Utilities.sleep(Math.min(wait, 60) * 1000);
      continue;
    }
    if (code < 200 || code >= 300) {
      throw new Error("multipart 요청 실패(" + code + "): " + text);
    }
    return JSON.parse(text);
  }
}

function buildMultipartPayload_(boundary, fields, fileBlob) {
  const chunks = [];
  const crlf = "\r\n";

  Object.keys(fields).forEach(function (k) {
    chunks.push(
      Utilities.newBlob(
        "--" + boundary + crlf +
          'Content-Disposition: form-data; name="' + k + '"' + crlf + crlf +
          String(fields[k]) + crlf
      ).getBytes()
    );
  });

  chunks.push(
    Utilities.newBlob(
      "--" + boundary + crlf +
        'Content-Disposition: form-data; name="file"; filename="' + fileBlob.getName() + '"' + crlf +
        "Content-Type: " + (fileBlob.getContentType() || "application/octet-stream") + crlf + crlf
    ).getBytes()
  );
  chunks.push(fileBlob.getBytes());
  chunks.push(Utilities.newBlob(crlf + "--" + boundary + "--" + crlf).getBytes());

  return flattenBytes_(chunks);
}

function flattenBytes_(arrays) {
  let total = 0;
  arrays.forEach(function (a) {
    total += a.length;
  });
  const out = new Uint8Array(total);
  let offset = 0;
  arrays.forEach(function (a) {
    out.set(a, offset);
    offset += a.length;
  });
  return out;
}

function getProp_(key, defVal) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === undefined || v === "") ? defVal : v;
}

function validateReq_(obj, requiredKeys) {
  requiredKeys.forEach(function (k) {
    if (obj[k] === undefined || obj[k] === null || String(obj[k]).trim() === "") {
      throw new Error("필수값 누락: " + k);
    }
  });
}

