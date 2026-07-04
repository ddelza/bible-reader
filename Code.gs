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
  return { book: book, chapter: chapterNum, verses: verses };
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
