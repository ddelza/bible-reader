const SS_ID = '1gkx1N02cdvMj5GHoUjiqXqlxE8aiuNyqqhao4jpYvuw';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const callback = e && e.parameter && e.parameter.callback;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('성경 읽기')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  let data;
  try {
    if (action === 'getBooks') data = getBooks();
    else if (action === 'getChapterVerses') data = getChapterVerses(e.parameter.book, e.parameter.chapter);
    else if (action === 'submitRecord') data = submitRecord(e.parameter.book, e.parameter.chapter, e.parameter.verse, e.parameter.date, e.parameter.author, e.parameter.content);
    else if (action === 'deleteRecord') data = deleteRecord(e.parameter.book, e.parameter.chapter, e.parameter.verse, e.parameter.ts);
    else if (action === 'getAuthors') data = getAuthors();
    else if (action === 'getPrayers') data = getPrayers();
    else if (action === 'submitPrayer') data = submitPrayer(e.parameter.type, e.parameter.author, e.parameter.date, e.parameter.content);
    else if (action === 'updatePrayer') data = updatePrayer(e.parameter.ts, e.parameter.type, e.parameter.author, e.parameter.content);
    else if (action === 'deletePrayer') data = deletePrayer(e.parameter.ts);
    else data = { error: 'unknown action: ' + action };
  } catch (err) {
    data = { error: err.toString() };
  }

  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// 성경명 시트에서 책 목록(번호, 이름, 장 수, 색인)을 가져온다.
function getBooks() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, ['성경명']);
  if (!sheet) return { error: '성경명 시트를 찾을 수 없습니다.' };

  const raw = sheet.getDataRange().getValues();
  const header = raw[0];
  const idxNum = findColIndex(header, ['번호']);
  const idxName = findColIndex(header, ['성경명']);
  const idxChapters = findColIndex(header, ['장 수', '장수']);
  const idxIndex = findColIndex(header, ['색인']);

  const books = [];
  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    const code = row[idxIndex];
    const num = row[idxNum];
    if (!code || !num) continue; // 마지막 합계 행 등 스킵

    const fullName = String(row[idxName] || '');
    const koName = fullName.split('(')[0].trim();
    books.push({
      번호: num,
      이름: koName || fullName,
      장수: row[idxChapters],
      색인: String(code).trim()
    });
  }
  return books;
}

// 성경 시트에서 특정 책(색인)의 특정 장에 속한 모든 절을 가져온다.
// 시트가 책→장→절 순으로 정렬되어 있다는 전제로, 전체를 스캔하지 않고
// 필요한 컬럼/행만 단계적으로 좁혀서 읽는다 (데이터가 3만 행 이상이라 성능상 중요).
function getChapterVerses(book, chapter) {
  if (!book || !chapter) return { error: '색인과 장이 필요합니다.' };
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, ['성경']);
  if (!sheet) return { error: '성경 시트를 찾을 수 없습니다.' };

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxBook = findColIndex(header, ['색인']);
  const idxChapter = findColIndex(header, ['장']);
  const idxVerse = findColIndex(header, ['절']);
  const idxGae = findColIndex(header, ['개역개정']);
  const idxSae = findColIndex(header, ['새번역']);
  const idxEsv = findColIndex(header, ['ESV']);

  const n = sheet.getLastRow() - 1;
  if (n <= 0) return { error: '데이터가 없습니다.' };

  // 1단계: 색인 컬럼만 읽어서 해당 책의 행 범위를 찾는다.
  const bookCol = sheet.getRange(2, idxBook + 1, n, 1).getValues();
  let bookStart = -1, bookEnd = -1;
  for (let i = 0; i < bookCol.length; i++) {
    if (String(bookCol[i][0]).trim() === book) {
      if (bookStart === -1) bookStart = i;
      bookEnd = i;
    } else if (bookStart !== -1) {
      break;
    }
  }
  if (bookStart === -1) return { error: '해당 색인을 찾을 수 없습니다: ' + book };

  const bookRowCount = bookEnd - bookStart + 1;
  const bookStartRow = 2 + bookStart;

  // 2단계: 그 책 범위 안에서 장 컬럼만 읽어서 해당 장의 행 범위를 찾는다.
  const chapterNum = parseInt(chapter, 10);
  const chapterCol = sheet.getRange(bookStartRow, idxChapter + 1, bookRowCount, 1).getValues();
  let chStart = -1, chEnd = -1;
  for (let i = 0; i < chapterCol.length; i++) {
    if (parseInt(chapterCol[i][0], 10) === chapterNum) {
      if (chStart === -1) chStart = i;
      chEnd = i;
    } else if (chStart !== -1) {
      break;
    }
  }
  if (chStart === -1) return { error: '해당 장을 찾을 수 없습니다: ' + chapter };

  const chRowCount = chEnd - chStart + 1;
  const chStartRow = bookStartRow + chStart;

  // 3단계: 절 + 번역본 컬럼만 읽는다.
  const minIdx = Math.min(idxVerse, idxGae, idxSae, idxEsv);
  const maxIdx = Math.max(idxVerse, idxGae, idxSae, idxEsv);
  const values = sheet.getRange(chStartRow, minIdx + 1, chRowCount, maxIdx - minIdx + 1).getValues();

  const verses = values.map(row => ({
    절: row[idxVerse - minIdx],
    개역개정: row[idxGae - minIdx],
    새번역: row[idxSae - minIdx],
    ESV: row[idxEsv - minIdx]
  }));

  verses.sort((a, b) => a.절 - b.절);

  // 각 절에 가족기록(개수/내용) 붙이기 (이미 열어둔 ss를 재사용 — 매번 openById를
  // 새로 하면 이 스프레드시트처럼 큰 파일에서는 그만큼 느려진다)
  const records = loadRecordsFor(ss, book, chapterNum);
  verses.forEach(v => {
    const list = records[v.절] || [];
    v.기록 = list;
    v.기록수 = list.length;
  });

  return { book: book, chapter: chapterNum, verses: verses };
}

// 성경명 시트의 작성자 목록(F2:F10)을 가져온다. 빈 칸은 제외한다.
function getAuthors() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = findSheet(ss, ['성경명']);
  if (!sheet) return { error: '성경명 시트를 찾을 수 없습니다.' };

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxAuthor = findColIndex(header, ['작성자']);
  if (idxAuthor === -1) return [];

  const values = sheet.getRange(2, idxAuthor + 1, 9, 1).getValues();
  return values.map(r => String(r[0] || '').trim()).filter(v => v);
}

// =============================================
// 가족 기록 (댓글처럼 절마다 누적되는 날짜/작성자/내용)
// =============================================

const RECORD_SHEET_NAME = '가족기록';
const RECORD_HEADER = ['색인', '장', '절', '날짜', '기록자', '내용', '등록시각'];

function getOrCreateRecordSheet(ss) {
  let sheet = ss.getSheetByName(RECORD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RECORD_SHEET_NAME);
    sheet.appendRow(RECORD_HEADER);
  }
  return sheet;
}

// 특정 책의 특정 장에 속한 모든 기록을 절 번호별로 묶어서 반환한다.
function loadRecordsFor(ss, book, chapter) {
  const sheet = getOrCreateRecordSheet(ss);
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < 2) return result;

  const data = sheet.getRange(2, 1, lastRow - 1, RECORD_HEADER.length).getValues();
  data.forEach(row => {
    const [rBook, rChapter, rVerse, date, author, content, ts] = row;
    if (String(rBook).trim() !== book) return;
    if (parseInt(rChapter, 10) !== parseInt(chapter, 10)) return;
    if (!author && !content) return;

    const verseNum = parseInt(rVerse, 10);
    if (!result[verseNum]) result[verseNum] = [];
    result[verseNum].push({
      날짜: formatRecordDate(date),
      기록자: String(author || ''),
      내용: String(content || ''),
      등록시각: ts instanceof Date ? ts.toISOString() : String(ts || '')
    });
  });

  Object.keys(result).forEach(k => {
    result[k].sort((a, b) => a.등록시각.localeCompare(b.등록시각));
  });
  return result;
}

function formatRecordDate(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d || '');
}

// 새 기록을 추가하고, 그 절의 최신 기록 목록을 반환한다 (프론트가 다시 조회하지 않아도 되도록).
function submitRecord(book, chapter, verse, date, author, content) {
  const chapterNum = parseInt(chapter, 10);
  const verseNum = parseInt(verse, 10);
  author = String(author || '').trim();
  content = String(content || '').trim();

  if (!book || !chapterNum || !verseNum || !author || !content) {
    return { error: '날짜, 작성자, 내용을 모두 입력해주세요.' };
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreateRecordSheet(ss);
  sheet.appendRow([book, chapterNum, verseNum, date, author, content, new Date()]);

  const records = loadRecordsFor(ss, book, chapterNum);
  return { records: records[verseNum] || [] };
}

// 등록시각(밀리초 단위 ISO 문자열, 사실상 고유 키)으로 특정 기록 한 건을 찾아 삭제한다.
function deleteRecord(book, chapter, verse, ts) {
  const chapterNum = parseInt(chapter, 10);
  const verseNum = parseInt(verse, 10);
  if (!book || !chapterNum || !verseNum || !ts) {
    return { error: '삭제할 기록을 특정할 수 없습니다.' };
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreateRecordSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, RECORD_HEADER.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowTs = row[6] instanceof Date ? row[6].toISOString() : String(row[6] || '');
      if (String(row[0]).trim() === book && parseInt(row[1], 10) === chapterNum &&
          parseInt(row[2], 10) === verseNum && rowTs === ts) {
        sheet.deleteRow(2 + i);
        break;
      }
    }
  }

  const records = loadRecordsFor(ss, book, chapterNum);
  return { records: records[verseNum] || [] };
}

// =============================================
// 기도제목
// =============================================

const PRAYER_SHEET_NAME = '기도제목';
const PRAYER_HEADER = ['종류', '작성자', '날짜', '내용', '등록시각'];

function getOrCreatePrayerSheet(ss) {
  let sheet = ss.getSheetByName(PRAYER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PRAYER_SHEET_NAME);
    sheet.appendRow(PRAYER_HEADER);
  }
  return sheet;
}

// 전체 기도제목을 최신 등록순으로 반환한다. 검색/작성자/기간 필터는 프론트에서
// 처리한다 — 성경 본문과 달리 데이터량이 적어 서버에서 범위를 좁힐 필요가 없다.
function getPrayers() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreatePrayerSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, PRAYER_HEADER.length).getValues();
  const prayers = data.map(row => {
    const [type, author, date, content, ts] = row;
    if (!author && !content) return null;
    return {
      종류: String(type || ''),
      작성자: String(author || ''),
      날짜: formatRecordDate(date),
      내용: String(content || ''),
      등록시각: ts instanceof Date ? ts.toISOString() : String(ts || '')
    };
  }).filter(Boolean);

  prayers.sort((a, b) => b.등록시각.localeCompare(a.등록시각));
  return prayers;
}

function submitPrayer(type, author, date, content) {
  type = String(type || '').trim();
  author = String(author || '').trim();
  content = String(content || '').trim();
  if (!type || !author || !content) {
    return { error: '종류, 작성자, 내용을 모두 입력해주세요.' };
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreatePrayerSheet(ss);
  sheet.appendRow([type, author, date, content, new Date()]);

  return getPrayers();
}

// 등록시각(밀리초 단위 ISO 문자열, 사실상 고유 키)으로 기도제목 한 건을 찾아 종류/작성자/
// 내용을 수정한다. 날짜와 등록시각(정렬·식별 키)은 그대로 둔다.
function updatePrayer(ts, type, author, content) {
  type = String(type || '').trim();
  author = String(author || '').trim();
  content = String(content || '').trim();
  if (!ts || !type || !author || !content) {
    return { error: '종류, 작성자, 내용을 모두 입력해주세요.' };
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreatePrayerSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, PRAYER_HEADER.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowTs = data[i][4] instanceof Date ? data[i][4].toISOString() : String(data[i][4] || '');
      if (rowTs === ts) {
        const row = 2 + i;
        sheet.getRange(row, 1, 1, 2).setValues([[type, author]]);
        sheet.getRange(row, 4, 1, 1).setValues([[content]]);
        break;
      }
    }
  }
  return getPrayers();
}

// 등록시각(밀리초 단위 ISO 문자열, 사실상 고유 키)으로 기도제목 한 건을 찾아 삭제한다.
function deletePrayer(ts) {
  if (!ts) return { error: '삭제할 기도제목을 특정할 수 없습니다.' };

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = getOrCreatePrayerSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, PRAYER_HEADER.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowTs = data[i][4] instanceof Date ? data[i][4].toISOString() : String(data[i][4] || '');
      if (rowTs === ts) {
        sheet.deleteRow(2 + i);
        break;
      }
    }
  }
  return getPrayers();
}

// =============================================
// 재사용 유틸리티
// =============================================

function findSheet(ss, candidates) {
  for (const name of candidates) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  const all = ss.getSheets();
  for (const cand of candidates) {
    const found = all.find(s => s.getName().includes(cand.replace(/\s/g, '')));
    if (found) return found;
  }
  return null;
}

function findColIndex(header, candidates) {
  for (const cand of candidates) {
    const idx = header.findIndex(h => String(h).replace(/\s/g, '').includes(cand.replace(/\s/g, '')));
    if (idx >= 0) return idx;
  }
  return -1;
}
