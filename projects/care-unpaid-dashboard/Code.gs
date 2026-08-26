/*************************************************
 * Bookkeeping Subscription SaaS — 미납 관리 대시보드
 *
 * 주간 미납 CSV 업로드 → 사업자번호 그룹화 → 미납현황·월별 탭 자동 적재
 * → 완납/강제해지 상태 자동 처리 → Chart.js 대시보드 → 슬랙 주간 리포트.
 *
 * 미납 회수율 86.23% (2026년 7월) 달성의 데이터 인프라.
 *************************************************/

const SHEET = {
  CHECKIN:   '체크인목록',
  INPUT:     '주간_입력',
  STATUS:    '미납현황',
  DASHBOARD: '대시보드',
  BLACKLIST: '블랙리스트',
};

const CK = {
  BIZ_NO:      2,
  MANAGER:     16,
  FEE:         9,
  STATE:       8,
  PAY_DATE:    17,
  PHONE:       4,
  CONTRACT_DT: 7,
};

const RAW = {
  BIZ_NO:   13,
  BIZ_NAME: 7,
  ORG_NAME: 12,
  REASON:   3,
  PRODUCT:  4,
  PAY_TYPE: 8,
  YEAR:     9,
  MONTH:    10,
  AMOUNT:   11,
  OWNER:    14,
  PHONE:    15,
};

const STATUS_HEADER = [
  '대표자명', '연락처', '상호', '사업자번호', '상태',
  '기장료', '미납액', '미납개월수', '사유', '상품',
  '결제수단', '담당자', '귀속월', '결제일', '최근확인일', '완납일'
];

// 상태(0) + 등록일(1) 추가, 사업자번호 → index 5
// 상태값: '예정' = 3개월+ 미납 현재 고객 / '완료' = 실제 강제해지된 고객
const BLACKLIST_HEADER = [
  '상태', '등록일', '대표자명', '연락처', '상호', '사업자번호',
  '기장료', '미납액', '미납개월수', '담당자', '비고'
];

// 웹 대시보드 배포 URL (Script Properties로 이동 권장)
const DASHBOARD_URL = PropertiesService.getScriptProperties().getProperty('DASHBOARD_URL') || '';

// =============================================
// 제외 명단 (이용 고객 아님) — 실제 사업자번호는 마스킹
// 원본 운영 시에는 5건의 제외 대상이 하드코딩되어 있었음
// =============================================
const EXCLUDE_BIZ_NO = new Set([
  // 'XXXXXXXXXX',  // 예시: 상호명 (대표자명) - 이용 고객 아님
]);

// 세일즈 담당자 제외 리스트 (Script Properties로 관리 권장)
const EXCLUDED_MANAGERS = ['EXCLUDED_MANAGER_A', 'EXCLUDED_MANAGER_B'];

// =============================================
// 고객 구간 계산
// =============================================
function getCustomerSegment(payDateRaw, contractDateRaw, today) {
  var dateRaw = payDateRaw || contractDateRaw;
  if (!dateRaw) return '미분류';
  var d = dateRaw instanceof Date ? dateRaw : new Date(String(dateRaw).trim());
  if (isNaN(d.getTime())) {
    if (contractDateRaw) {
      d = contractDateRaw instanceof Date ? contractDateRaw : new Date(String(contractDateRaw).trim());
      if (isNaN(d.getTime())) return '미분류';
    } else return '미분류';
  }
  var months = (today.getFullYear() - d.getFullYear()) * 12 + (today.getMonth() - d.getMonth());
  if (months < 2)  return '초기 (2개월 미만)';
  if (months < 6)  return '중기 (2~6개월)';
  return '장기 (6개월 이상)';
}

// =============================================
// 귀속월 계산
// =============================================
function calcAttrMonth(ymList, today) {
  if (ymList && ymList.length > 0) {
    return ymList[ymList.length - 1];
  }
  return Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM');
}

// =============================================
// 메인 실행 함수
// =============================================
function syncUnpaidData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt('CSV 뽑은 날짜 입력', 'CSV를 뽑은 날짜를 입력하세요 (예: 2026-02-04)\n오늘 날짜로 하려면 그냥 확인을 누르세요.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.CANCEL) return;

  var inputDate = response.getResponseText().trim();
  var today;
  if (inputDate === '') {
    today = new Date();
  } else {
    today = new Date(inputDate);
    if (isNaN(today.getTime())) {
      ui.alert('❌ 날짜 형식이 올바르지 않아요. yyyy-MM-dd 형식으로 입력해 주세요.');
      return;
    }
  }

  var todayStr  = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM-dd');
  var thisMonth = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM');

  var checkinSheet = ss.getSheetByName(SHEET.CHECKIN);
  var inputSheet   = ss.getSheetByName(SHEET.INPUT);
  var statusSheet  = ss.getSheetByName(SHEET.STATUS);

  if (!checkinSheet || !inputSheet || !statusSheet) {
    SpreadsheetApp.getUi().alert('❌ 탭 이름 확인:\n체크인목록 / 주간_입력 / 미납현황');
    return;
  }

  // 1. 체크인목록 → 모수 맵
  var checkinRows    = checkinSheet.getDataRange().getValues().slice(1);
  var masterMap      = {};
  var totalCount     = 0;
  var canceledBizSet = new Set();

  checkinRows.forEach(function(r) {
    var state   = String(r[CK.STATE]).trim();
    var manager = String(r[CK.MANAGER]).trim();
    var fee     = parseInt(String(r[CK.FEE]).replace(/[^0-9]/g, '')) || 0;
    var bizNo   = String(r[CK.BIZ_NO]).trim();
    var phone   = String(r[CK.PHONE]).trim();

    if (EXCLUDE_BIZ_NO.has(bizNo)) return;
    if (state === '기장해지-취소') { canceledBizSet.add(bizNo); return; }
    if (state !== '기장중' || fee < 11000 || EXCLUDED_MANAGERS.indexOf(manager) !== -1) return;

    var payDay     = '';
    var payDateRaw = r[CK.PAY_DATE];
    if (payDateRaw) {
      var d = payDateRaw instanceof Date ? payDateRaw : new Date(String(payDateRaw).trim());
      if (!isNaN(d.getTime())) payDay = d.getDate();
    }

    var segment = getCustomerSegment(r[CK.PAY_DATE], r[CK.CONTRACT_DT], today);
    masterMap[bizNo] = { manager: manager, payDay: payDay, phone: phone, fee: fee, segment: segment };
    totalCount++;
  });

  Logger.log('모수: ' + totalCount + ' / 기장해지-취소: ' + canceledBizSet.size);

  // 2. 원본 CSV 로드 → 사업자번호 기준 그룹화
  var inputRows = inputSheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[RAW.BIZ_NO]).trim() !== ''; });

  var grouped = {};
  inputRows.forEach(function(r) {
    var bizNo = String(r[RAW.BIZ_NO]).trim();
    if (!masterMap[bizNo]) return;
    if (EXCLUDE_BIZ_NO.has(bizNo)) return;

    var ownerCheck = String(r[RAW.OWNER] || '').trim();
    if (!ownerCheck) return;

    var year    = String(r[RAW.YEAR]).trim();
    var month   = String(r[RAW.MONTH]).trim();
    var amount  = Number(r[RAW.AMOUNT]) || 0;
    var product = String(r[RAW.PRODUCT]).trim();

    if (!year || year === '-' || !month || month === '-' || amount <= 0) return;
    if (product.indexOf('종소세') !== -1) return;

    var ym = year + '-' + month.padStart(2, '0');
    if (!grouped[bizNo]) grouped[bizNo] = { rows: [], latestYM: ym, latestRow: r };
    grouped[bizNo].rows.push({ row: r, ym: ym });
    if (ym > grouped[bizNo].latestYM) { grouped[bizNo].latestYM = ym; grouped[bizNo].latestRow = r; }
  });

  Logger.log('그룹화된 사업자 수: ' + Object.keys(grouped).length);

  // 3. 헤더 설정
  statusSheet.getRange(1, 1, 1, STATUS_HEADER.length).setValues([STATUS_HEADER]);
  statusSheet.getRange(1, 1, 1, STATUS_HEADER.length).setFontWeight('bold');
  statusSheet.setFrozenRows(1);
  statusSheet.getRange('N:N').setNumberFormat('@');

  // 4. 기존 데이터 로드 (강제해지 감지용)
  var prevStatusData = statusSheet.getDataRange().getValues().slice(1);
  var prevStatusMap  = {};
  prevStatusData.forEach(function(r) {
    var bizNo = String(r[3]).trim();
    if (bizNo) prevStatusMap[bizNo] = r;
  });

  // 5. 기존 데이터 삭제
  var lastRow = statusSheet.getLastRow();
  if (lastRow > 1) {
    statusSheet.getRange(2, 1, lastRow - 1, STATUS_HEADER.length).clearContent();
    statusSheet.getRange(2, 1, lastRow - 1, STATUS_HEADER.length).setBackground(null);
  }

  // 6. 새 데이터 배열 생성
  var newRows = [];
  Object.keys(grouped).forEach(function(bizNo) {
    var data      = grouped[bizNo];
    var master    = masterMap[bizNo];
    var latestRow = data.latestRow;
    var ymList    = data.rows.map(function(d) { return d.ym; }).sort();
    var count     = ymList.length;
    var reason    = String(latestRow[RAW.REASON]   || '').trim();
    var product   = String(latestRow[RAW.PRODUCT]  || '').trim();
    var payType   = String(latestRow[RAW.PAY_TYPE] || '').trim();
    var owner     = String(latestRow[RAW.OWNER]    || '').trim();
    var phone     = master.phone || String(latestRow[RAW.PHONE] || '').trim();
    var bizName   = String(latestRow[RAW.BIZ_NAME] || latestRow[RAW.ORG_NAME] || '').trim();
    var fee       = master.fee;
    var attrMonth = calcAttrMonth(ymList, today);
    var payDayStr = master.payDay ? master.payDay + '일' : '';

    newRows.push([
      owner,           // 0. 대표자명
      phone,           // 1. 연락처
      bizName,         // 2. 상호
      bizNo,           // 3. 사업자번호
      '미납',          // 4. 상태
      fee,             // 5. 기장료
      fee * count,     // 6. 미납액
      count,           // 7. 미납개월수
      reason,          // 8. 사유
      product,         // 9. 상품
      payType,         // 10. 결제수단
      master.manager,  // 11. 담당자
      attrMonth,       // 12. 귀속월
      payDayStr,       // 13. 결제일
      todayStr,        // 14. 최근확인일
      ''               // 15. 완납일
    ]);
  });

  Logger.log('newRows: ' + newRows.length);

  // 7. 미납현황 쓰기
  if (newRows.length > 0) {
    statusSheet.getRange(2, 1, newRows.length, STATUS_HEADER.length).setValues(newRows);
  }

  // 8. 강제해지 감지 → recordBlacklist (완료 처리, 예정→완료 승격)
  var forceTerminated = [];
  var allSheets = ss.getSheets();
  allSheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (!/^\d{4}-\d{2}$/.test(name)) return;
    var sheetData = sheet.getDataRange().getValues().slice(1);
    sheetData.forEach(function(r) {
      var bizNo  = String(r[3]).trim();
      var status = String(r[4]).trim();
      var months = Number(r[7]) || 0;
      if (canceledBizSet.has(bizNo) && status === '미납' && months >= 3) {
        var alreadyAdded = forceTerminated.some(function(f) { return f.bizNo === bizNo; });
        if (!alreadyAdded) forceTerminated.push({ bizNo: bizNo, row: r });
      }
    });
  });
  if (forceTerminated.length > 0) recordBlacklist(ss, forceTerminated, todayStr);

  // 강제해지 예정자 upsert (flush 후 미납현황 탭에서 직접 읽음)
  SpreadsheetApp.flush();
  updateForceTerminationList(ss, todayStr);

  // 9. currentBizSet 구성
  var currentBizSet = new Set();
  inputRows.forEach(function(r) {
    var bizNo = String(r[RAW.BIZ_NO]).trim();
    if (bizNo) currentBizSet.add(bizNo);
  });
  newRows.forEach(function(r) { currentBizSet.add(String(r[3]).trim()); });
  var currentAttrMonthMap = {};
  newRows.forEach(function(r) {
    var bizNo = String(r[3]).trim();
    currentAttrMonthMap[bizNo] = String(r[12]);
  });

  updateMonthlySheets(ss, newRows, todayStr, currentBizSet, canceledBizSet, prevStatusMap, currentAttrMonthMap);

  // 10. 대시보드 업데이트
  updateDashboardSheet(ss, newRows, todayStr, thisMonth, totalCount, masterMap);

  var msg = '✅ 업데이트 완료! (' + newRows.length + '건)';
  if (forceTerminated.length > 0) msg += '\n🚨 강제해지 감지: ' + forceTerminated.length + '건';
  SpreadsheetApp.getUi().alert(msg);
}

// =============================================
// 월별 탭 적재 + 완납/강제해지 처리
// =============================================
function updateMonthlySheets(ss, newRows, todayStr, currentBizSet, canceledBizSet, prevStatusMap, currentAttrMonthMap) {
  var monthDataMap = {};
  newRows.forEach(function(row) {
    var attrMonth = String(row[12]).trim();
    if (!/^\d{4}-\d{2}$/.test(attrMonth)) return;
    if (parseInt(attrMonth.split('-')[0]) < 2026) return;
    if (!monthDataMap[attrMonth]) monthDataMap[attrMonth] = {};
    monthDataMap[attrMonth][String(row[3])] = row;
  });

  Object.keys(monthDataMap).forEach(function(ym) {
    var bizMap = monthDataMap[ym];
    var sheet  = ss.getSheetByName(ym);
    if (!sheet) sheet = ss.insertSheet(ym);

    sheet.getRange(1, 1, 1, STATUS_HEADER.length).setValues([STATUS_HEADER]);
    sheet.getRange(1, 1, 1, STATUS_HEADER.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('N:N').setNumberFormat('@');

    var existingData = sheet.getDataRange().getValues().slice(1);
    var existingMap  = {};
    existingData.forEach(function(r, i) { existingMap[String(r[3]).trim()] = i + 2; });

    var toAppend = [];
    Object.keys(bizMap).forEach(function(bizNo) {
      var row     = bizMap[bizNo];
      var rowData = row.slice(0, STATUS_HEADER.length);
      while (rowData.length < STATUS_HEADER.length) rowData.push('');
      if (existingMap[bizNo]) {
        sheet.getRange(existingMap[bizNo], 1, 1, STATUS_HEADER.length).setValues([rowData]);
        sheet.getRange(existingMap[bizNo], 1, 1, STATUS_HEADER.length).setBackground(null);
      } else {
        toAppend.push(rowData);
      }
    });
    if (toAppend.length > 0) {
      var lr = sheet.getLastRow();
      sheet.getRange(lr + 1, 1, toAppend.length, STATUS_HEADER.length).setValues(toAppend);
    }

    var updatedData = sheet.getDataRange().getValues().slice(1);
    updatedData.forEach(function(r, i) {
      var bizNo      = String(r[3]).trim();
      var status     = String(r[4]).trim();
      var rowNum     = i + 2;
      var prevMonths = Number(r[7]) || 0;

      // 이 월의 귀속월이 아닌 고객(다른 월로 이동했거나 완납한 고객)도 완납 처리
      // 이전 버전은 CSV에서 완전히 사라진 고객만 완납 처리해서 잔여 미납 버그가 있었음
      if (!bizMap[bizNo] && status === '미납') {
        var currentAttrMonth = currentAttrMonthMap ? (currentAttrMonthMap[bizNo] || '') : '';
        if (currentAttrMonth && currentAttrMonth > ym) {
          return; // 미납 누적 중 - 건드리지 않음
        }
        if (canceledBizSet.has(bizNo) && prevMonths >= 3) {
          sheet.getRange(rowNum, 5).setValue('강제해지');
          sheet.getRange(rowNum, 16).setValue(todayStr);
          sheet.getRange(rowNum, 1, 1, STATUS_HEADER.length).setBackground('#f4c7c3');
        } else {
          sheet.getRange(rowNum, 5).setValue('완납');
          sheet.getRange(rowNum, 16).setValue(todayStr);
          sheet.getRange(rowNum, 1, 1, STATUS_HEADER.length).setBackground('#d9f7d9');
        }
      }
    });
  });
}

// =============================================
// 블랙리스트 완료 기록
// (강제해지 확정 → 완료 처리, 예정 행 있으면 승격)
// =============================================
function recordBlacklist(ss, forceTerminated, todayStr) {
  var sheet = ss.getSheetByName(SHEET.BLACKLIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.BLACKLIST);
    sheet.getRange(1, 1, 1, BLACKLIST_HEADER.length).setValues([BLACKLIST_HEADER]);
    sheet.getRange(1, 1, 1, BLACKLIST_HEADER.length).setFontWeight('bold').setBackground('#fce8e8');
    sheet.setFrozenRows(1);
  }

  var existing    = sheet.getDataRange().getValues().slice(1);
  var existingMap = {};
  existing.forEach(function(r, i) {
    var bizNo = String(r[5]).trim();
    if (bizNo) existingMap[bizNo] = { rowNum: i + 2, status: String(r[0]).trim() };
  });

  var toAppend = [];
  forceTerminated.forEach(function(item) {
    var r     = item.row;
    var bizNo = item.bizNo;

    if (existingMap[bizNo]) {
      if (existingMap[bizNo].status === '완료') return;
      // 예정 → 완료 승격
      var rowNum = existingMap[bizNo].rowNum;
      sheet.getRange(rowNum, 1).setValue('완료');
      sheet.getRange(rowNum, 2).setValue(todayStr);
      sheet.getRange(rowNum, 11).setValue('미납으로 인한 강제해지');
      sheet.getRange(rowNum, 1, 1, BLACKLIST_HEADER.length).setBackground('#fff0f0');
    } else {
      toAppend.push(['완료', todayStr, r[0], r[1], r[2], r[3], r[5], r[6], r[7], r[11], '미납으로 인한 강제해지']);
    }
  });

  if (toAppend.length > 0) {
    var lr = sheet.getLastRow();
    sheet.getRange(lr + 1, 1, toAppend.length, BLACKLIST_HEADER.length).setValues(toAppend);
    sheet.getRange(lr + 1, 1, toAppend.length, BLACKLIST_HEADER.length).setBackground('#fff0f0');
  }
}

// =============================================
// 강제해지 예정자 upsert
// (3개월+ 미납 → 예정 / 해소 시 자동 제거 / 완료 행은 절대 건드리지 않음)
// =============================================
function updateForceTerminationList(ss, todayStr) {
  var sheet = ss.getSheetByName(SHEET.BLACKLIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.BLACKLIST);
    sheet.getRange(1, 1, 1, BLACKLIST_HEADER.length).setValues([BLACKLIST_HEADER]);
    sheet.getRange(1, 1, 1, BLACKLIST_HEADER.length).setFontWeight('bold').setBackground('#fce8e8');
    sheet.setFrozenRows(1);
  }

  var statusSheet = ss.getSheetByName(SHEET.STATUS);
  var currentForceMap = {};
  statusSheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (String(row[4]).trim() === '미납' && (Number(row[7]) || 0) >= 3) {
      currentForceMap[String(row[3]).trim()] = row;
    }
  });

  var existing = sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[5]).trim() !== ''; });

  var resultRows   = [];
  var processedSet = new Set();

  existing.forEach(function(r) {
    var status = String(r[0]).trim();
    var bizNo  = String(r[5]).trim();

    if (status === '완료') {
      resultRows.push(r.slice(0, BLACKLIST_HEADER.length));
      processedSet.add(bizNo);

    } else if (status === '예정') {
      if (currentForceMap[bizNo]) {
        // 여전히 3개월+ → 최신 미납 수치로 업데이트, 등록일은 원래 날짜 유지
        var row = currentForceMap[bizNo];
        resultRows.push([
          '예정',
          r[1],        // 등록일 유지
          row[0], row[1], row[2], row[3],
          row[5], row[6], row[7], row[11],
          r[10] || ''
        ]);
        processedSet.add(bizNo);
      }
      // else: 3개월 미만으로 해소 → 행 drop (자동 제거)
    }
  });

  // 신규 예정자 추가
  Object.keys(currentForceMap).forEach(function(bizNo) {
    if (processedSet.has(bizNo)) return;
    var row = currentForceMap[bizNo];
    resultRows.push([
      '예정', todayStr,
      row[0], row[1], row[2], row[3],
      row[5], row[6], row[7], row[11], ''
    ]);
  });

  // 시트 재작성 (완료 위, 예정 아래)
  resultRows.sort(function(a, b) {
    var order = { '완료': 0, '예정': 1 };
    return (order[a[0]] || 9) - (order[b[0]] || 9);
  });

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, BLACKLIST_HEADER.length).clearContent();
    sheet.getRange(2, 1, lastRow - 1, BLACKLIST_HEADER.length).setBackground(null);
  }

  if (resultRows.length > 0) {
    sheet.getRange(2, 1, resultRows.length, BLACKLIST_HEADER.length).setValues(resultRows);
    resultRows.forEach(function(r, i) {
      var bg = String(r[0]).trim() === '완료' ? '#fff0f0' : '#fff8e1';
      sheet.getRange(i + 2, 1, 1, BLACKLIST_HEADER.length).setBackground(bg);
    });
  }
}

// =============================================
// 대시보드 탭 업데이트
// =============================================
function updateDashboardSheet(ss, newRows, todayStr, thisMonth, totalCount, masterMap) {
  var dash  = ss.getSheetByName(SHEET.DASHBOARD);
  var today = new Date();
  if (!dash) { dash = ss.insertSheet(SHEET.DASHBOARD); initDashboard(dash); }

  var sheets       = ss.getSheets();
  var monthlyStats = {};
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (!/^\d{4}-\d{2}$/.test(name)) return;
    var sheetData = sheet.getDataRange().getValues().slice(1)
      .filter(function(r) { return String(r[3]).trim() !== ''; });

    var segCount = { '초기 (2개월 미만)': 0, '중기 (2~6개월)': 0, '장기 (6개월 이상)': 0 };
    sheetData.forEach(function(r) {
      var status = String(r[4]).trim();
      if (status !== '미납') return;
      var bizNo = String(r[3]).trim();
      var seg   = (masterMap[bizNo] && masterMap[bizNo].segment) || '미분류';
      if (segCount[seg] !== undefined) segCount[seg]++;
    });

    monthlyStats[name] = {
      total:  sheetData.length,
      unpaid: sheetData.filter(function(r) { return String(r[4]).trim() === '미납'; }).length,
      paid:   sheetData.filter(function(r) { return String(r[4]).trim() === '완납'; }).length,
      force:  sheetData.filter(function(r) { return String(r[4]).trim() === '강제해지'; }).length,
      seg:    segCount,
    };
  });

  var ymList      = Object.keys(monthlyStats).sort();
  var allDashData = dash.getDataRange().getValues();
  var dataStartRow = -1;
  for (var i = 0; i < allDashData.length; i++) {
    if (String(allDashData[i][0]).trim() === '귀속월') { dataStartRow = i + 2; break; }
  }
  if (dataStartRow === -1) { initDashboard(dash); allDashData = dash.getDataRange().getValues(); dataStartRow = 7; }

  var existingYMMap = {};
  for (var i = dataStartRow - 1; i < allDashData.length; i++) {
    var ymRaw = allDashData[i][0];
    var ym    = ymRaw instanceof Date ? Utilities.formatDate(ymRaw, 'Asia/Seoul', 'yyyy-MM') : String(ymRaw).trim();
    if (ym && /^\d{4}-\d{2}$/.test(ym)) {
      if (!existingYMMap[ym]) { existingYMMap[ym] = i + 1; if (ymRaw instanceof Date) dash.getRange(i+1,1).setValue(ym); }
      else { dash.deleteRow(i + 1); allDashData = dash.getDataRange().getValues(); i--; }
    }
  }

  ymList.forEach(function(ym) {
    if (!existingYMMap[ym]) { var lr = dash.getLastRow(); dash.getRange(lr+1,1).setValue(ym); existingYMMap[ym] = lr+1; }
  });

  var updatedData = dash.getDataRange().getValues();
  for (var row = dataStartRow - 1; row < updatedData.length; row++) {
    var ymRaw2 = updatedData[row][0];
    var ym2    = ymRaw2 instanceof Date ? Utilities.formatDate(ymRaw2, 'Asia/Seoul', 'yyyy-MM') : String(ymRaw2).trim();
    if (!ym2 || !/^\d{4}-\d{2}$/.test(ym2)) continue;
    var stats = monthlyStats[ym2];
    if (!stats) continue;
    var rowNum = row + 1;

    var closeDate = String(updatedData[row][6] || '').trim();
    if (!closeDate) {
      var parts      = ym2.split('-');
      var closeMonth = parseInt(parts[1]) + 1;
      var closeYear  = parseInt(parts[0]);
      if (closeMonth > 12) { closeMonth = 1; closeYear++; }
      var closeEndDate = new Date(closeYear, closeMonth, 0);
      if (today >= closeEndDate) {
        closeDate = Utilities.formatDate(closeEndDate, 'Asia/Seoul', 'yyyy-MM-dd');
        dash.getRange(rowNum, 7).setValue(closeDate);
      }
    }

    dash.getRange(rowNum, 3).setValue(stats.total);
    if (!closeDate) dash.getRange(rowNum, 5).setValue(stats.unpaid);
    dash.getRange(rowNum, 8).setValue(stats.paid);
    dash.getRange(rowNum, 9).setValue(stats.force);
    dash.getRange(rowNum, 4).setFormula('=IFERROR(C'+rowNum+'/B'+rowNum+',"")');
    dash.getRange(rowNum, 6).setFormula('=IFERROR(E'+rowNum+'/B'+rowNum+',"")');
    dash.getRange(rowNum, 10).setFormula('=IFERROR(H'+rowNum+'/C'+rowNum+',"")');
  }

  var segTotalMap  = {};
  var segUnpaidMap = {};
  Object.keys(masterMap).forEach(function(bizNo) {
    var seg = masterMap[bizNo].segment || '미분류';
    segTotalMap[seg] = (segTotalMap[seg] || 0) + 1;
  });
  newRows.forEach(function(row) {
    var seg = masterMap[String(row[3])] ? masterMap[String(row[3])].segment : '미분류';
    segUnpaidMap[seg] = (segUnpaidMap[seg] || 0) + 1;
  });

  var sc       = 16;
  var segments = ['초기 (2개월 미만)', '중기 (2~6개월)', '장기 (6개월 이상)', '미분류'];

  dash.getRange(6, sc, 20, 4).clearContent();
  dash.getRange(6, sc, 20, 4).setBackground(null);

  dash.getRange(6, sc, 1, 4).setValues([['고객구간', '모수', '미납건수', '미납률']]);
  dash.getRange(6, sc, 1, 4).setFontWeight('bold').setBackground('#534AB7').setFontColor('#ffffff');

  var segRows = segments.map(function(seg) {
    var total  = segTotalMap[seg]  || 0;
    var unpaid = segUnpaidMap[seg] || 0;
    return [seg, total, unpaid, total > 0 ? Math.round(unpaid/total*10000)/100+'%' : '-'];
  });
  dash.getRange(7, sc, segRows.length, 4).setValues(segRows);
}

// =============================================
// 대시보드 초기화
// =============================================
function initDashboard(dash) {
  dash.clearContents();
  dash.clearFormats();
  dash.getRange('A1').setValue('미납 현황 대시보드').setFontSize(18).setFontWeight('bold');
  dash.getRange('A2').setValue('※ 모수(B열)는 매월 직접 입력하세요.').setFontSize(11).setFontColor('#888780');
  dash.getRange('A3').setValue('※ 종결일(G열)은 귀속월 +1개월 말일에 자동 입력됩니다.').setFontSize(11).setFontColor('#888780');
  dash.getRange('A4').setValue('※ 최초/최종미납수, 완납수, 강제해지수는 월별 탭에서 자동 계산됩니다.').setFontSize(11).setFontColor('#888780');
  var header = ['귀속월','모수','최초미납수','최초미납률','최종미납수','최종미납률','종결일','완납수','강제해지수','회수율','비고','초기미납수(2개월미만)','중기미납수(2~6개월)','장기미납수(6개월이상)'];
  dash.getRange('A6:N6').setValues([header]);
  dash.getRange('A6:N6').setFontWeight('bold').setBackground('#1D9E75').setFontColor('#ffffff');
  dash.setFrozenRows(6);
  var widths = [90,80,90,90,90,90,100,80,90,80,120,130,130,130];
  widths.forEach(function(w,i) { dash.setColumnWidth(i+1, w); });
  dash.getRange('D7:D200').setNumberFormat('0.00%');
  dash.getRange('F7:F200').setNumberFormat('0.00%');
  dash.getRange('J7:J200').setNumberFormat('0.00%');
}

// =============================================
// 고객 이력 조회
// =============================================
function searchCustomer() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('고객조회');
  if (!sheet) {
    sheet = ss.insertSheet('고객조회');
    initCustomerSearch(sheet);
  }

  var query = String(sheet.getRange('A2').getValue()).trim();
  if (!query) {
    SpreadsheetApp.getUi().alert('A2셀에 사업자번호 또는 이름을 입력해 주세요!');
    return;
  }

  sheet.getRange('A5:F100').clearContent();
  sheet.getRange('A5:F100').setBackground(null);

  var checkinSheet = ss.getSheetByName(SHEET.CHECKIN);
  var checkinRows  = checkinSheet.getDataRange().getValues().slice(1);
  var customer     = null;
  checkinRows.forEach(function(r) {
    var bizNo = String(r[CK.BIZ_NO]).trim();
    var name  = String(r[3] || '').trim();
    var biz   = String(r[1] || '').trim();
    if (bizNo === query || name === query || biz.indexOf(query) !== -1) {
      customer = r;
    }
  });

  if (!customer) {
    sheet.getRange('A5').setValue('❌ 해당 고객을 찾을 수 없어요.');
    sheet.getRange('A5').setFontColor('#E24B4A').setFontWeight('bold');
    return;
  }

  var bizNo   = String(customer[CK.BIZ_NO]).trim();
  var name    = String(customer[3] || '').trim();
  var biz     = String(customer[1] || '').trim();
  var phone   = String(customer[CK.PHONE] || '').trim();
  var manager = String(customer[CK.MANAGER] || '').trim();
  var state   = String(customer[CK.STATE] || '').trim();
  var segment = getCustomerSegment(customer[CK.PAY_DATE], customer[CK.CONTRACT_DT], new Date());

  sheet.getRange('A5').setValue('👤 고객 기본 정보').setFontWeight('bold').setBackground('#f1efe8');
  sheet.getRange('A6').setValue('대표자명');   sheet.getRange('B6').setValue(name);
  sheet.getRange('A7').setValue('상호');      sheet.getRange('B7').setValue(biz);
  sheet.getRange('A8').setValue('사업자번호'); sheet.getRange('B8').setValue(bizNo);
  sheet.getRange('A9').setValue('연락처');    sheet.getRange('B9').setValue(phone);
  sheet.getRange('A10').setValue('담당자');   sheet.getRange('B10').setValue(manager);
  sheet.getRange('A11').setValue('기장상태');  sheet.getRange('B11').setValue(state);
  sheet.getRange('A12').setValue('고객구간');  sheet.getRange('B12').setValue(segment);

  var statusSheet = ss.getSheetByName(SHEET.STATUS);
  var statusRows  = statusSheet.getDataRange().getValues().slice(1);
  var currentRow  = null;
  statusRows.forEach(function(r) { if (String(r[3]).trim() === bizNo) currentRow = r; });

  sheet.getRange('A14').setValue('💰 현재 미납 현황').setFontWeight('bold').setBackground('#f1efe8');
  if (currentRow) {
    sheet.getRange('A15').setValue('상태');      sheet.getRange('B15').setValue(currentRow[4]);
    sheet.getRange('A16').setValue('미납개월수'); sheet.getRange('B16').setValue(currentRow[7] + '개월');
    sheet.getRange('A17').setValue('미납액');    sheet.getRange('B17').setValue(currentRow[6]);
    sheet.getRange('B17').setNumberFormat('#,##0"원"');
    sheet.getRange('A18').setValue('미납 사유'); sheet.getRange('B18').setValue(currentRow[8]);
    sheet.getRange('A19').setValue('결제수단');  sheet.getRange('B19').setValue(currentRow[10]);
    sheet.getRange('A20').setValue('귀속월');   sheet.getRange('B20').setValue(currentRow[12]);
  } else {
    sheet.getRange('A15').setValue('현재 미납 내역 없음').setFontColor('#1D9E75');
  }

  sheet.getRange('A22').setValue('📅 월별 미납 이력').setFontWeight('bold').setBackground('#f1efe8');
  sheet.getRange('A23:F23').setValues([['귀속월', '상태', '미납액', '미납개월수', '사유', '완납일']]);
  sheet.getRange('A23:F23').setFontWeight('bold').setBackground('#e8f4fd');

  var sheets  = ss.getSheets();
  var history = [];
  sheets.forEach(function(s) {
    var n = s.getName();
    if (!/^\d{4}-\d{2}$/.test(n)) return;
    var data = s.getDataRange().getValues().slice(1);
    data.forEach(function(r) {
      if (String(r[3]).trim() === bizNo) {
        history.push([n, r[4], r[6], r[7], r[8], r[15] || '-']);
      }
    });
  });
  history.sort(function(a, b) { return a[0].localeCompare(b[0]); });

  if (history.length > 0) {
    sheet.getRange(24, 1, history.length, 6).setValues(history);
    sheet.getRange(24, 3, history.length, 1).setNumberFormat('#,##0"원"');
    history.forEach(function(h, i) {
      var bg = null;
      if (h[1] === '완납') bg = '#d9f7d9';
      else if (h[1] === '강제해지') bg = '#f4c7c3';
      if (bg) sheet.getRange(24 + i, 1, 1, 6).setBackground(bg);
    });
  } else {
    sheet.getRange('A24').setValue('미납 이력 없음').setFontColor('#1D9E75');
  }

  var blacklistSheet = ss.getSheetByName(SHEET.BLACKLIST);
  var isBlacklisted  = false;
  var blStatus       = '';
  if (blacklistSheet) {
    var blRows = blacklistSheet.getDataRange().getValues().slice(1);
    blRows.forEach(function(r) {
      if (String(r[5]).trim() === bizNo) {
        isBlacklisted = true;
        blStatus      = String(r[0]).trim();
      }
    });
  }
  if (isBlacklisted) {
    var blackRow = 24 + history.length + 2;
    var label, bg, color;
    if (blStatus === '예정') {
      label = '⚠️ 강제해지 예정자 (3개월+ 미납)';
      bg    = '#FFF8E1';
      color = '#E65100';
    } else {
      label = '⚫ 강제해지 완료 (블랙리스트 등록)';
      bg    = '#FCEBEB';
      color = '#A32D2D';
    }
    sheet.getRange(blackRow, 1).setValue(label).setFontWeight('bold').setFontColor(color).setBackground(bg);
  }
}

// =============================================
// 고객조회 탭 초기화
// =============================================
function initCustomerSearch(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange('A1').setValue('🔍 고객 이력 조회').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('').setBackground('#fffbcc');
  sheet.getRange('A2').setNote('사업자번호 또는 대표자명을 입력 후\n[미납관리] → [고객 조회] 메뉴 클릭');
  sheet.getRange('A3').setValue('※ 사업자번호 or 대표자명 입력 후 [미납관리 → 고객 조회] 클릭').setFontColor('#888780').setFontSize(11);
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
}

function openDashboard() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + DASHBOARD_URL + '", "_blank"); google.script.host.close();<\/script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, '대시보드 열기');
}

// =============================================
// 웹앱 (Chart.js 대시보드)
// =============================================
function doGet() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var statusSheet  = ss.getSheetByName(SHEET.STATUS);
  var checkinSheet = ss.getSheetByName(SHEET.CHECKIN);
  var dashSheet    = ss.getSheetByName(SHEET.DASHBOARD);
  var today        = new Date();
  var todayStr     = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM-dd');
  var thisMonth    = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM');

  var data = statusSheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[4]).trim() === '미납'; })
    .map(function(r) {
      var am = r[12];
      if (am instanceof Date) r[12] = Utilities.formatDate(am, 'Asia/Seoul', 'yyyy-MM');
      else r[12] = String(am).trim();
      return r;
    });

  var checkinRows = checkinSheet.getDataRange().getValues().slice(1);
  var totalCount  = 0;
  var segTotalMap = {};
  checkinRows.forEach(function(r) {
    var state   = String(r[CK.STATE]).trim();
    var manager = String(r[CK.MANAGER]).trim();
    var fee     = parseInt(String(r[CK.FEE]).replace(/[^0-9]/g, '')) || 0;
    var bizNo   = String(r[CK.BIZ_NO]).trim();
    if (EXCLUDE_BIZ_NO.has(bizNo)) return;
    if (state !== '기장중' || fee < 11000 || EXCLUDED_MANAGERS.indexOf(manager) !== -1) return;
    totalCount++;
    var seg = getCustomerSegment(r[CK.PAY_DATE], r[CK.CONTRACT_DT], today);
    segTotalMap[seg] = (segTotalMap[seg] || 0) + 1;
  });

  var totalUnpaid = data.reduce(function(s,r) { return s+(Number(r[6])||0); }, 0);
  var unpaidCount = data.length;
  var forceCount  = data.filter(function(r) { return (Number(r[7])||0) >= 3; }).length;
  var unpaidRate  = totalCount > 0 ? Math.round(unpaidCount/totalCount*1000)/10 : 0;

  var monthMap = {};
  data.forEach(function(row) {
    var ym = String(row[12]);
    if (!monthMap[ym]) monthMap[ym] = { count:0, amount:0, force:0 };
    monthMap[ym].count++;
    monthMap[ym].amount += Number(row[6])||0;
    if ((Number(row[7])||0) >= 3) monthMap[ym].force++;
  });

  var forceList = data.filter(function(r) { return (Number(r[7])||0) >= 3; })
    .sort(function(a,b) { return (Number(b[7])||0)-(Number(a[7])||0); })
    .map(function(r) { return { name:r[0], phone:r[1], biz:r[2], bizNo:r[3], months:r[7], amount:r[6], manager:r[11] }; });

  var monthRows = Object.keys(monthMap).sort().map(function(m) {
    var v = monthMap[m];
    return { ym:m, count:v.count, amount:v.amount, force:v.force, rate: totalCount > 0 ? Math.round(v.count/totalCount*1000)/10 : 0 };
  });

  var trendRows = [];
  if (dashSheet) {
    var dashData = dashSheet.getDataRange().getValues();
    var started  = false;
    dashData.forEach(function(r) {
      if (String(r[0]).trim() === '귀속월') { started = true; return; }
      if (!started) return;
      var ym = r[0] instanceof Date ? Utilities.formatDate(r[0],'Asia/Seoul','yyyy-MM') : String(r[0]).trim();
      if (!ym || !/^\d{4}-\d{2}$/.test(ym)) return;
      trendRows.push({
        ym: ym, mosu: Number(r[1])||0,
        firstUnpaid: Number(r[2])||0, firstRate: typeof r[3]==='number'?r[3]:0,
        finalUnpaid: Number(r[4])||0, finalRate: typeof r[5]==='number'?r[5]:0,
        paid: Number(r[7])||0, force: Number(r[8])||0,
        recoverRate: typeof r[9]==='number'?r[9]:0, note: String(r[10]||''),
        segEarly: Number(r[11])||0, segMid: Number(r[12])||0, segLate: Number(r[13])||0,
      });
    });
  }

  var segUnpaidMap = {};
  data.forEach(function(row) {
    var bizNo = String(row[3]).trim();
    var seg   = '미분류';
    checkinRows.forEach(function(r) {
      if (String(r[CK.BIZ_NO]).trim() === bizNo) {
        seg = getCustomerSegment(r[CK.PAY_DATE], r[CK.CONTRACT_DT], today);
      }
    });
    segUnpaidMap[seg] = (segUnpaidMap[seg] || 0) + 1;
  });
  var segmentRows = ['초기 (2개월 미만)','중기 (2~6개월)','장기 (6개월 이상)'].map(function(seg) {
    var total  = segTotalMap[seg]  || 0;
    var unpaid = segUnpaidMap[seg] || 0;
    return { seg:seg, total:total, unpaid:unpaid, rate: total>0?Math.round(unpaid/total*1000)/10:0 };
  });

  var tmpl = HtmlService.createTemplate(getHtmlTemplate());
  tmpl.data = JSON.stringify({
    today:todayStr, thisMonth:thisMonth,
    totalCount:totalCount, unpaidCount:unpaidCount,
    unpaidRate:unpaidRate, totalUnpaid:totalUnpaid,
    forceCount:forceCount, monthRows:monthRows,
    forceList:forceList, trendRows:trendRows,
    segmentRows:segmentRows,
  });
  return tmpl.evaluate().setTitle('미납 현황 대시보드').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtmlTemplate() {
  return '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>미납 현황 대시보드</title><script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"><\/script><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f5f4f0;color:#1a1a18;padding:2rem}h1{font-size:22px;font-weight:500;margin-bottom:4px}.sub{font-size:13px;color:#888780;margin-bottom:2rem}.stat-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:2rem}.stat{background:#fff;border-radius:12px;border:.5px solid rgba(0,0,0,.08);padding:1rem 1.25rem}.stat-label{font-size:12px;color:#888780;margin-bottom:6px}.stat-val{font-size:24px;font-weight:500}.stat-val.red{color:#E24B4A}.card{background:#fff;border-radius:12px;border:.5px solid rgba(0,0,0,.08);padding:1.25rem;margin-bottom:1.5rem}.card-title{font-size:15px;font-weight:500;margin-bottom:1rem}table{width:100%;border-collapse:collapse;font-size:13px}th{text-align:left;padding:8px 10px;background:#f5f4f0;color:#5f5e5a;font-weight:500;border-bottom:.5px solid rgba(0,0,0,.08)}td{padding:8px 10px;border-bottom:.5px solid rgba(0,0,0,.05)}tr:last-child td{border-bottom:none}.force-row{background:#fff5f5}.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:500}.badge.red{background:#FCEBEB;color:#A32D2D}.chart-wrap{position:relative;height:260px}.two-col{display:grid;grid-template-columns:1fr 1fr;gap:1.5rem;margin-bottom:1.5rem}@media(max-width:700px){.stat-grid{grid-template-columns:repeat(2,1fr)}.two-col{grid-template-columns:1fr}}<\/style><\/head><body><h1>미납 현황 대시보드<\/h1><div class="sub" id="updated"><\/div><div class="stat-grid"><div class="stat"><div class="stat-label">전체 모수<\/div><div class="stat-val" id="totalCount"><\/div><\/div><div class="stat"><div class="stat-label">미납 건수<\/div><div class="stat-val red" id="unpaidCount"><\/div><\/div><div class="stat"><div class="stat-label">미납률<\/div><div class="stat-val red" id="unpaidRate"><\/div><\/div><div class="stat"><div class="stat-label">총 미납금액<\/div><div class="stat-val red" id="totalUnpaid"><\/div><\/div><div class="stat"><div class="stat-label">강제해지 대상<\/div><div class="stat-val red" id="forceCount"><\/div><\/div><\/div><div class="card"><div class="card-title">월별 미납 추이 (최초 vs 최종)<\/div><div class="chart-wrap"><canvas id="trendChart"><\/canvas><\/div><\/div><div class="card"><div class="card-title">월별 상세 현황<\/div><table><thead><tr><th>귀속월<\/th><th>모수<\/th><th>최초미납<\/th><th>최초미납률<\/th><th>최종미납<\/th><th>최종미납률<\/th><th>완납<\/th><th>회수율<\/th><\/tr><\/thead><tbody id="trendTable"><\/tbody><\/table><\/div><div class="card"><div class="card-title">고객 구간별 월별 미납 추이<\/div><div class="chart-wrap"><canvas id="segChart"><\/canvas><\/div><\/div><div class="card"><div class="card-title">고객 구간별 상세<\/div><table><thead><tr><th>귀속월<\/th><th>초기 (2개월 미만)<\/th><th>중기 (2~6개월)<\/th><th>장기 (6개월 이상)<\/th><th>합계<\/th><\/tr><\/thead><tbody id="segMonthTable"><\/tbody><\/table><\/div><div class="card"><div class="card-title">강제해지 대상 (3개월↑)<\/div><table><thead><tr><th>대표자명<\/th><th>연락처<\/th><th>상호<\/th><th>미납개월<\/th><th>미납액<\/th><th>담당자<\/th><\/tr><\/thead><tbody id="forceTable"><\/tbody><\/table><\/div><script>const d=JSON.parse(\'<?= data ?>\');const fmt=v=>Number(v).toLocaleString("ko-KR")+"원";const pct=v=>typeof v==="number"?(v*100).toFixed(2)+"%":"-";document.getElementById("updated").textContent="기준일: "+d.today;document.getElementById("totalCount").textContent=d.totalCount.toLocaleString();document.getElementById("unpaidCount").textContent=d.unpaidCount.toLocaleString();document.getElementById("unpaidRate").textContent=d.unpaidRate+"%";document.getElementById("totalUnpaid").textContent=fmt(d.totalUnpaid);document.getElementById("forceCount").textContent=d.forceCount.toLocaleString();if(d.trendRows&&d.trendRows.length>0){new Chart(document.getElementById("trendChart"),{type:"bar",data:{labels:d.trendRows.map(r=>r.ym),datasets:[{label:"최초미납수",data:d.trendRows.map(r=>r.firstUnpaid),backgroundColor:"#E24B4A66",yAxisID:"y"},{label:"최종미납수",data:d.trendRows.map(r=>r.finalUnpaid),backgroundColor:"#1D9E7566",yAxisID:"y"},{label:"회수율(%)",data:d.trendRows.map(r=>typeof r.recoverRate==="number"?Math.round(r.recoverRate*1000)/10:null),type:"line",borderColor:"#534AB7",backgroundColor:"transparent",yAxisID:"y1",tension:0.3,spanGaps:true}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{font:{size:12}}}},scales:{y:{position:"left",title:{display:true,text:"건수"}},y1:{position:"right",title:{display:true,text:"회수율(%)"},min:0,max:100,grid:{drawOnChartArea:false}}}}});new Chart(document.getElementById("segChart"),{type:"bar",data:{labels:d.trendRows.map(r=>r.ym),datasets:[{label:"초기 (2개월 미만)",data:d.trendRows.map(r=>r.segEarly),backgroundColor:"#E24B4A99"},{label:"중기 (2~6개월)",data:d.trendRows.map(r=>r.segMid),backgroundColor:"#378ADD99"},{label:"장기 (6개월 이상)",data:d.trendRows.map(r=>r.segLate),backgroundColor:"#1D9E7599"}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{font:{size:12}}}},scales:{y:{title:{display:true,text:"미납건수"}}}}});const tt=document.getElementById("trendTable");d.trendRows.forEach(r=>{tt.innerHTML+="<tr><td>"+r.ym+"<\/td><td>"+(r.mosu||"-")+"<\/td><td>"+(r.firstUnpaid||"-")+"<\/td><td>"+pct(r.firstRate)+"<\/td><td>"+(r.finalUnpaid||"-")+"<\/td><td>"+pct(r.finalRate)+"<\/td><td>"+(r.paid||"-")+"<\/td><td>"+pct(r.recoverRate)+"<\/td><\/tr>";});const smt=document.getElementById("segMonthTable");d.trendRows.forEach(r=>{const total=r.segEarly+r.segMid+r.segLate;smt.innerHTML+="<tr><td>"+r.ym+"<\/td><td>"+r.segEarly+"건<\/td><td>"+r.segMid+"건<\/td><td>"+r.segLate+"건<\/td><td><strong>"+total+"건<\/strong><\/td><\/tr>";});}const ft=document.getElementById("forceTable");d.forceList.forEach(r=>{ft.innerHTML+="<tr class=\'force-row\'><td>"+r.name+"<\/td><td>"+r.phone+"<\/td><td>"+r.biz+"<\/td><td><span class=\'badge red\'>"+r.months+"개월<\/span><\/td><td>"+fmt(r.amount)+"<\/td><td>"+r.manager+"<\/td><\/tr>";});<\/script><\/body><\/html>';
}

// =============================================
// 메뉴 생성
// =============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('미납관리')
    .addItem('지금 업데이트', 'syncUnpaidData')
    .addToUi();
  ui.createMenu('📊 대시보드')
    .addItem('대시보드 열기', 'openDashboard')
    .addToUi();
}
