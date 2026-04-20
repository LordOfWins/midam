/**
 * 미담사진관 손님용 태블릿 웹 백엔드
 * Google Apps Script Web App - doGet/doPost JSON API
 *
 * v1.3.0 변경사항 (2026.04.20 3차 피드백):
 * - SHEET_ID 정정 (운영 시트로 교체)
 * - 전화번호 검증 완화: 4 / 7 / 8 / 11 자리만 허용
 * - 전화번호 컬럼에 setNumberFormat('@') 강제 적용 (앞자리 0 누락 방지)
 * - updateEntry 저장 시에도 텍스트 서식 재적용
 * - verifyPhone 허용 길이 완화 (4/7/8/11)
 *
 * v1.2.0 변경사항:
 * - appendRow 대신 ID 컬럼 기반 실제 마지막 데이터 행 탐색 후 직접 setValues
 *   (AppSheet가 빈 서식만 남긴 빈 행들 사이에 끼어드는 현상 해결)
 * - 한국 휴대폰 번호 엄격 검증 (010 + 11자리)
 * - 저장된 번호 11자리 아닌 경우 STORED_PHONE_CORRUPTED 에러 반환
 */

// ============================================================
// 설정값
// ============================================================
const CONFIG = {
  SHEET_ID: '11kcBZRYG1aqNn9qJEUqhHILy2sFTOSywZVOnUeLyi5k',
  SHEET_NAME: '미담_앱접수',
  API_TOKEN: 'midam-2026-secret-token',
  DEFAULT_STATUS: '촬영',
  EXCLUDE_STATUS: ['완료', '취소']
}

const COLUMNS = ['ID', '날짜', '상품', '상황', '이름', '전화번호', '이메일', '파일명', '인증키']

// ============================================================
// 엔트리포인트
// ============================================================

function doGet(e) {
  return handleRequest(e, 'GET')
}

function doPost(e) {
  return handleRequest(e, 'POST')
}

function handleRequest(e, method) {
  try {
    let params = {}
    let action = ''

    if (method === 'GET') {
      params = e.parameter || {}
      action = params.action || ''
    } else {
      if (e.postData && e.postData.contents) {
        params = JSON.parse(e.postData.contents)
        action = params.action || ''
      }
    }

    if (params.token !== CONFIG.API_TOKEN) {
      return jsonResponse({ ok: false, error: 'UNAUTHORIZED' })
    }

    switch (action) {
      case 'list':
        return jsonResponse(listWaiting())
      case 'create':
        return jsonResponse(createEntry(params.data))
      case 'verify':
        return jsonResponse(verifyPhone(params.id, params.last4))
      case 'update':
        return jsonResponse(updateEntry(params.id, params.data, params.last4))
      case 'ping':
        return jsonResponse({ ok: true, version: '1.3.0', time: new Date().toISOString() })
      default:
        return jsonResponse({ ok: false, error: 'UNKNOWN_ACTION' })
    }
  } catch (err) {
    Logger.log('handleRequest error: ' + err.stack)
    return jsonResponse({ ok: false, error: 'SERVER_ERROR', message: String(err) })
  }
}

// ============================================================
// 응답 헬퍼
// ============================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

// ============================================================
// 시트 접근 헬퍼
// ============================================================

function getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID)
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME)
  if (!sheet) {
    throw new Error('시트를 찾을 수 없습니다: ' + CONFIG.SHEET_NAME)
  }
  return sheet
}

function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn()
  if (lastCol === 0) {
    throw new Error('시트에 헤더가 없습니다')
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
  const map = {}
  headers.forEach((h, idx) => {
    map[String(h).trim()] = idx
  })
  return { map: map, headers: headers, lastCol: lastCol }
}

/**
 * ID 컬럼에 실제 값이 있는 마지막 행 번호를 반환 (1-based)
 * AppSheet가 남긴 빈 서식/포맷 행을 무시하고 실제 데이터 기준으로 판단
 *
 * 반환: 실제 데이터가 있는 마지막 행 번호 (없으면 1=헤더만 있음)
 */
function findLastDataRow(sheet) {
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return 1   // 헤더만 있거나 빈 시트

  const { map } = getHeaderMap(sheet)
  const idIdx = map['ID']
  if (idIdx === undefined) throw new Error('ID 컬럼을 찾을 수 없습니다')

  // ID 컬럼 전체 값을 한 번에 읽어서 뒤에서부터 탐색 (성능)
  const idColValues = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues()

  for (let i = idColValues.length - 1; i >= 0; i--) {
    const value = String(idColValues[i][0] || '').trim()
    if (value !== '') {
      return i + 2   // 0-index + 헤더(1행) 보정
    }
  }
  return 1   // 모두 비어있음 - 헤더 다음 행부터 시작
}

// ============================================================
// 전화번호 유효성 검증
// ============================================================

/**
 * 전화번호 검증 - 4 / 7 / 8 / 11 자리만 허용
 *
 * 허용 케이스:
 * - 4자리: 끝번호만 ("1234")
 * - 7자리: 010 + 끝번호 4자리 ("0101234")
 * - 8자리: 중간+끝 ("12345678")
 * - 11자리: 010 + 전체 ("01012345678")
 */
function validateFlexiblePhone(phone) {
  const digits = String(phone || '').replace(/\D/g, '')

  if (digits.length === 0) return { ok: false, error: 'PHONE_REQUIRED' }

  const allowed = [4, 7, 8, 11]
  if (allowed.indexOf(digits.length) === -1) {
    return { ok: false, error: 'PHONE_INVALID_LENGTH' }
  }

  // 7자리 / 11자리는 반드시 010으로 시작해야 함
  if ((digits.length === 7 || digits.length === 11) && !digits.startsWith('010')) {
    return { ok: false, error: 'PHONE_INVALID_PREFIX' }
  }

  return { ok: true, digits: digits }
}

/**
 * 길이별 하이픈 자동 포맷 (저장용)
 * - 4자리:  "1234"        -> "1234"
 * - 7자리:  "0101234"     -> "010-1234"
 * - 8자리:  "12345678"    -> "1234-5678"
 * - 11자리: "01012345678" -> "010-1234-5678"
 */
function normalizeFlexiblePhone(phone) {
  const digits = String(phone || '').replace(/\D/g, '')

  if (digits.length === 11) {
    return digits.slice(0, 3) + '-' + digits.slice(3, 7) + '-' + digits.slice(7, 11)
  }
  if (digits.length === 8) {
    return digits.slice(0, 4) + '-' + digits.slice(4, 8)
  }
  if (digits.length === 7) {
    return digits.slice(0, 3) + '-' + digits.slice(3, 7)
  }
  return digits  // 4자리는 그대로
}

// ============================================================
// 액션 1 - 대기리스트 조회
// ============================================================

function listWaiting() {
  const sheet = getSheet()
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) {
    return { ok: true, list: [] }
  }

  const { map, lastCol } = getHeaderMap(sheet)
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()

  const idIdx = map['ID']
  const nameIdx = map['이름']
  const statusIdx = map['상황']
  const dateIdx = map['날짜']

  if (idIdx === undefined || nameIdx === undefined || statusIdx === undefined) {
    throw new Error('필수 컬럼 누락 (ID / 이름 / 상황)')
  }

  const list = []
  for (let i = 0; i < data.length; i++) {
    const row = data[i]
    const status = String(row[statusIdx] || '').trim()

    if (CONFIG.EXCLUDE_STATUS.indexOf(status) !== -1) continue

    const id = String(row[idIdx] || '').trim()
    const name = String(row[nameIdx] || '').trim()
    if (!id || !name) continue

    list.push({
      id: id,
      name: name,
      date: formatDateYYMMDD(row[dateIdx])
    })
  }

  list.reverse()

  return { ok: true, list: list, count: list.length }
}

// ============================================================
// 액션 2 - 신규 접수 등록
// ============================================================

function createEntry(data) {
  if (!data || typeof data !== 'object') {
    return { ok: false, error: 'INVALID_DATA' }
  }

  const name = sanitize(data.name)
  const phone = sanitize(data.phone)
  const email = sanitize(data.email || '')

  if (!name) return { ok: false, error: 'NAME_REQUIRED' }

  const phoneCheck = validateFlexiblePhone(phone)
  if (!phoneCheck.ok) return phoneCheck

  const normalizedPhone = normalizeFlexiblePhone(phone)

  const lock = LockService.getScriptLock()
  try {
    lock.waitLock(10000)

    const sheet = getSheet()
    const { map, headers, lastCol } = getHeaderMap(sheet)

    const newRow = new Array(headers.length).fill('')
    const id = generateAppSheetCompatibleId()
    const today = formatDateYYMMDD(new Date())

    if (map['ID'] !== undefined) newRow[map['ID']] = id
    if (map['날짜'] !== undefined) newRow[map['날짜']] = today
    if (map['상품'] !== undefined) newRow[map['상품']] = ''
    if (map['상황'] !== undefined) newRow[map['상황']] = CONFIG.DEFAULT_STATUS
    if (map['이름'] !== undefined) newRow[map['이름']] = name
    if (map['전화번호'] !== undefined) newRow[map['전화번호']] = normalizedPhone
    if (map['이메일'] !== undefined) newRow[map['이메일']] = email
    if (map['파일명'] !== undefined) newRow[map['파일명']] = ''
    if (map['인증키'] !== undefined) newRow[map['인증키']] = ''

    // appendRow 대신 실제 데이터 기준 다음 행에 직접 삽입
    // (AppSheet가 남긴 빈 서식 행 무시)
    const targetRow = findLastDataRow(sheet) + 1

    // ⚠️ 반드시 setValues 호출 전에 텍스트 서식(@)을 지정해야 앞자리 0 보존됨
    // - 전화번호 컬럼: 숫자로 해석되지 않도록 텍스트 강제 (수정 3 핵심)
    // - ID 컬럼: 영숫자 혼합 ID 안전 보존
    if (map['전화번호'] !== undefined) {
      sheet.getRange(targetRow, map['전화번호'] + 1).setNumberFormat('@')
    }
    if (map['ID'] !== undefined) {
      sheet.getRange(targetRow, map['ID'] + 1).setNumberFormat('@')
    }

    sheet.getRange(targetRow, 1, 1, lastCol).setValues([newRow])

    return { ok: true, id: id, name: name, date: today, rowIndex: targetRow }
  } catch (err) {
    Logger.log('createEntry error: ' + err.stack)
    return { ok: false, error: 'CREATE_FAILED', message: String(err) }
  } finally {
    try { lock.releaseLock() } catch (e) {}
  }
}

// ============================================================
// 액션 3 - 전화번호 끝4자리 인증
// ============================================================

function verifyPhone(id, last4) {
  if (!id || !last4) return { ok: false, error: 'INVALID_PARAMS' }

  const last4Digits = String(last4).replace(/\D/g, '')
  if (last4Digits.length !== 4) return { ok: false, error: 'LAST4_INVALID' }

  const row = findRowById(id)
  if (!row) return { ok: false, error: 'NOT_FOUND' }

  const phoneDigits = String(row.data['전화번호'] || '').replace(/\D/g, '')

  // 허용 길이: 4 / 7 / 8 / 11 (수정 2 정책 반영)
  const allowedLengths = [4, 7, 8, 11]
  if (allowedLengths.indexOf(phoneDigits.length) === -1) {
    Logger.log('verifyPhone: 비정상 저장 번호 id=' + id + ' digits=[' + phoneDigits + '] length=' + phoneDigits.length)
    return { ok: false, error: 'STORED_PHONE_CORRUPTED', debug: phoneDigits.length }
  }

  // 4자리 저장 케이스는 번호 전체가 끝4자리
  // 그 외 케이스는 마지막 4자리로 비교
  const actualLast4 = phoneDigits.length === 4 ? phoneDigits : phoneDigits.slice(-4)

  if (actualLast4 !== last4Digits) {
    return { ok: false, error: 'LAST4_MISMATCH' }
  }

  return {
    ok: true,
    id: id,
    name: row.data['이름'] || '',
    phone: row.data['전화번호'] || '',
    email: row.data['이메일'] || ''
  }
}

// ============================================================
// 액션 4 - 정보 수정
// ============================================================

function updateEntry(id, data, last4) {
  if (!id || !data || !last4) return { ok: false, error: 'INVALID_PARAMS' }

  const verifyResult = verifyPhone(id, last4)
  if (!verifyResult.ok) return verifyResult

  const phone = sanitize(data.phone)
  const email = sanitize(data.email || '')

  const phoneCheck = validateFlexiblePhone(phone)
  if (!phoneCheck.ok) return phoneCheck

  const normalizedPhone = normalizeFlexiblePhone(phone)

  const lock = LockService.getScriptLock()
  try {
    lock.waitLock(10000)

    const sheet = getSheet()
    const { map } = getHeaderMap(sheet)
    const row = findRowById(id)
    if (!row) return { ok: false, error: 'NOT_FOUND' }

    if (map['전화번호'] !== undefined) {
      const phoneCell = sheet.getRange(row.rowIndex, map['전화번호'] + 1)
      phoneCell.setNumberFormat('@')  // 수정 시에도 텍스트 서식 재적용 (수정 3 핵심)
      phoneCell.setValue(normalizedPhone)
    }
    if (map['이메일'] !== undefined) {
      sheet.getRange(row.rowIndex, map['이메일'] + 1).setValue(email)
    }

    return { ok: true, id: id }
  } catch (err) {
    Logger.log('updateEntry error: ' + err.stack)
    return { ok: false, error: 'UPDATE_FAILED', message: String(err) }
  } finally {
    try { lock.releaseLock() } catch (e) {}
  }
}

// ============================================================
// 헬퍼 - ID로 행 찾기
// ============================================================

function findRowById(id) {
  const sheet = getSheet()
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return null

  const { map, headers, lastCol } = getHeaderMap(sheet)
  const idIdx = map['ID']
  if (idIdx === undefined) return null

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const targetId = String(id).trim()

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === targetId) {
      const rowObj = {}
      headers.forEach((h, idx) => {
        rowObj[String(h).trim()] = data[i][idx]
      })
      return {
        rowIndex: i + 2,
        data: rowObj
      }
    }
  }
  return null
}

// ============================================================
// 헬퍼 - 입력값 sanitize
// ============================================================

function sanitize(value) {
  if (value === null || value === undefined) return ''
  return String(value)
    .trim()
    .replace(/[\x00-\x1F\x7F]/g, '')
    .slice(0, 200)
}

// ============================================================
// 헬퍼 - AppSheet 호환 ID 생성
// ============================================================

function generateAppSheetCompatibleId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
  let id = ''
  for (let i = 0; i < 8; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length))
  }
  return id
}

// ============================================================
// 헬퍼 - 날짜 포맷 (YY-MM-DD)
// ============================================================

function formatDateYYMMDD(value) {
  if (!value) return ''

  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (/^\d{2}-\d{2}-\d{2}$/.test(trimmed)) return trimmed
    const parsed = new Date(trimmed)
    if (isNaN(parsed.getTime())) return trimmed
    return dateToYYMMDD(parsed)
  }

  if (value instanceof Date) {
    return dateToYYMMDD(value)
  }

  return String(value)
}

function dateToYYMMDD(date) {
  const yy = String(date.getFullYear()).slice(-2)
  const mm = String(date.getMonth() + 1).padStart(2, '0')
  const dd = String(date.getDate()).padStart(2, '0')
  return `${yy}-${mm}-${dd}`
}

// ============================================================
// 개발/운영 유틸
// ============================================================

function testList() {
  Logger.log(JSON.stringify(listWaiting(), null, 2))
}

function testCreate() {
  Logger.log(JSON.stringify(createEntry({
    name: '테스트',
    phone: '010-1234-5678',
    email: 'test@test.com'
  }), null, 2))
}

function testFindLastDataRow() {
  const sheet = getSheet()
  const result = findLastDataRow(sheet)
  Logger.log('실제 마지막 데이터 행: ' + result)
  Logger.log('sheet.getLastRow(): ' + sheet.getLastRow())
}

function testInspectRow() {
  const targetId = 'uFgqcZ2S'   // 문제 있는 ID로 교체해서 사용
  const row = findRowById(targetId)
  if (!row) {
    Logger.log('행 없음: ' + targetId)
    return
  }
  Logger.log('rowIndex: ' + row.rowIndex)
  Logger.log('전화번호 원본: [' + row.data['전화번호'] + ']')
  const digits = String(row.data['전화번호'] || '').replace(/\D/g, '')
  Logger.log('숫자만: [' + digits + '] length=' + digits.length)
  Logger.log('끝4자리: [' + digits.slice(-4) + ']')
}

/**
 * ⚠️ 1회성 실행 유틸 - 전화번호 컬럼 전체에 텍스트 서식(@) 적용
 *
 * 실행 시점:
 * - v1.3.0 배포 직후 1회 실행
 * - 이후 신규 등록분은 createEntry에서 자동 처리됨
 *
 * 동작:
 * - 전화번호/ID 컬럼 전체(헤더 포함)에 setNumberFormat('@') 적용
 * - 이미 저장된 값의 표시 방식이 텍스트로 바뀜
 * - ⚠️ 이미 숫자로 저장되어 앞자리 0이 사라진 값은 복구되지 않음
 *   -> 해당 행은 AppSheet나 시트에서 수동으로 하이픈 포함 형태로 재입력 필요
 *
 * 실행 방법:
 * 1. Apps Script 에디터 함수 드롭다운 -> applyTextFormatToPhoneColumn 선택
 * 2. 실행
 * 3. 실행 로그(Ctrl+Enter)에서 결과 확인
 */
function applyTextFormatToPhoneColumn() {
  const sheet = getSheet()
  const { map } = getHeaderMap(sheet)
  const maxRows = sheet.getMaxRows()

  const targets = ['전화번호', 'ID']
  const applied = []

  for (const col of targets) {
    const idx = map[col]
    if (idx === undefined) {
      Logger.log('⚠️ 컬럼 없음: ' + col)
      continue
    }
    // 전체 컬럼(헤더 포함 ~ 시트 끝 행)에 텍스트 서식 적용
    sheet.getRange(1, idx + 1, maxRows, 1).setNumberFormat('@')
    applied.push(col + ' (컬럼 ' + (idx + 1) + ')')
  }

  Logger.log('✅ 텍스트 서식 적용 완료: ' + applied.join(', '))
  Logger.log('⚠️ 이미 앞자리 0이 사라진 데이터는 수동 재입력 필요')
}

/**
 * ⚠️ 진단 유틸 - 전화번호 컬럼에서 비정상 길이 데이터 탐지
 *
 * 출력: 숫자만 추출 시 4/7/8/11 외의 길이를 가진 행 전부
 * -> 이 행들은 v1.3.0 배포 전에 저장되어 앞자리 0이 사라진 것일 가능성 높음
 */
function diagnoseCorruptedPhones() {
  const sheet = getSheet()
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) {
    Logger.log('데이터 없음')
    return
  }

  const { map, lastCol } = getHeaderMap(sheet)
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()

  const idIdx = map['ID']
  const phoneIdx = map['전화번호']
  const nameIdx = map['이름']

  if (phoneIdx === undefined) {
    Logger.log('전화번호 컬럼 없음')
    return
  }

  const allowed = [4, 7, 8, 11]
  const corrupted = []

  for (let i = 0; i < data.length; i++) {
    const rawPhone = String(data[i][phoneIdx] || '')
    const digits = rawPhone.replace(/\D/g, '')

    if (digits.length === 0) continue  // 빈 값은 무시
    if (allowed.indexOf(digits.length) !== -1) continue  // 정상

    corrupted.push({
      rowIndex: i + 2,
      id: String(data[i][idIdx] || ''),
      name: String(data[i][nameIdx] || ''),
      raw: rawPhone,
      digits: digits,
      length: digits.length
    })
  }

  Logger.log('=== 비정상 전화번호 진단 결과 ===')
  Logger.log('전체 검사 행: ' + data.length)
  Logger.log('비정상 행: ' + corrupted.length)
  Logger.log('')

  if (corrupted.length === 0) {
    Logger.log('✅ 모든 데이터 정상')
    return
  }

  for (const c of corrupted) {
    Logger.log('행 ' + c.rowIndex + ' | ID=' + c.id + ' | 이름=' + c.name +
               ' | 원본=[' + c.raw + '] | 숫자=[' + c.digits + '] length=' + c.length)
  }
  Logger.log('')
  Logger.log('⚠️ 위 행들은 AppSheet 또는 시트에서 수동 수정 필요')
  Logger.log('   (앞자리 0이 이미 사라진 경우 원본 번호를 알 수 없음)')
}

/**
 * ⚠️ 운영 유틸 - 빈 행 정리
 * 데이터 행(ID 값 있는 행) 사이사이 및 마지막 뒤쪽에 있는 빈 서식 행을 삭제
 *
 * 실행 방법:
 * 1. 함수 드롭다운 -> cleanupEmptyRows 선택 -> 실행
 * 2. 반드시 시트 백업 후 실행 (Apps Script 에디터에서 수동 확인 후 실행)
 *
 * 동작:
 * - findLastDataRow() 뒤쪽에 있는 모든 빈 행 삭제
 * - 데이터 행 사이에 낀 빈 행은 건드리지 않음 (AppSheet 레코드 삭제 흔적일 수 있음)
 */
function cleanupEmptyRows() {
  const sheet = getSheet()
  const lastRow = sheet.getLastRow()
  const lastDataRow = findLastDataRow(sheet)

  Logger.log('현재 sheet.getLastRow(): ' + lastRow)
  Logger.log('실제 마지막 데이터 행: ' + lastDataRow)

  if (lastRow <= lastDataRow) {
    Logger.log('정리할 빈 행 없음')
    return
  }

  const rowsToDelete = lastRow - lastDataRow
  Logger.log('삭제할 빈 행 수: ' + rowsToDelete + ' (' + (lastDataRow + 1) + '행~' + lastRow + '행)')

  // 실제 삭제 (⚠️ 실행 시 되돌릴 수 없음)
  sheet.deleteRows(lastDataRow + 1, rowsToDelete)

  Logger.log('✅ 정리 완료. 새 getLastRow(): ' + sheet.getLastRow())
}

/**
 * ⚠️ 실행 전 반드시 구글시트 백업 필수!
 *
 * 데이터 행 사이에 낀 빈 행들을 압축(삭제)하는 유틸
 *
 * 동작 방식:
 * 1. 모든 행을 훑으면서 ID 컬럼에 값이 있는지 확인
 * 2. ID가 빈 행들을 찾아서 삭제 목록 생성
 * 3. 뒤에서부터 삭제 (앞에서 삭제하면 행 번호가 밀려서 오류)
 *
 * 안전장치:
 * - dryRun 모드 기본값 true (실제 삭제 안 하고 로그만)
 * - false로 바꿔야 실제 삭제 실행
 * - 삭제 전 최종 삭제 대상 행 수를 로그에 출력
 */
function compactEmptyRows() {
  const DRY_RUN = true   // ⚠️ 실제 삭제하려면 false로 변경 후 재실행

  const sheet = getSheet()
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) {
    Logger.log('데이터 없음 - 정리할 행 없음')
    return
  }

  const { map } = getHeaderMap(sheet)
  const idIdx = map['ID']
  if (idIdx === undefined) throw new Error('ID 컬럼 없음')

  // 전체 ID 컬럼 값 로드 (2행 ~ lastRow)
  const idValues = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues()

  // 빈 행 번호 수집 (실제 시트 기준 1-based)
  const emptyRowNumbers = []
  for (let i = 0; i < idValues.length; i++) {
    const value = String(idValues[i][0] || '').trim()
    if (value === '') {
      emptyRowNumbers.push(i + 2)   // 0-index + 헤더 보정
    }
  }

  Logger.log('=== compactEmptyRows 진단 ===')
  Logger.log('DRY_RUN 모드: ' + DRY_RUN)
  Logger.log('전체 행 수 (데이터 영역): ' + (lastRow - 1))
  Logger.log('빈 행 개수: ' + emptyRowNumbers.length)
  Logger.log('데이터 행 개수: ' + (lastRow - 1 - emptyRowNumbers.length))

  if (emptyRowNumbers.length === 0) {
    Logger.log('삭제할 빈 행 없음 - 종료')
    return
  }

  // 빈 행의 연속 구간 요약 (로그 가독성)
  Logger.log('')
  Logger.log('--- 빈 행 연속 구간 ---')
  let rangeStart = emptyRowNumbers[0]
  let rangeEnd = emptyRowNumbers[0]
  for (let i = 1; i < emptyRowNumbers.length; i++) {
    if (emptyRowNumbers[i] === rangeEnd + 1) {
      rangeEnd = emptyRowNumbers[i]
    } else {
      Logger.log(rangeStart + '행 ~ ' + rangeEnd + '행 (' + (rangeEnd - rangeStart + 1) + '행)')
      rangeStart = emptyRowNumbers[i]
      rangeEnd = emptyRowNumbers[i]
    }
  }
  Logger.log(rangeStart + '행 ~ ' + rangeEnd + '행 (' + (rangeEnd - rangeStart + 1) + '행)')

  if (DRY_RUN) {
    Logger.log('')
    Logger.log('⚠️ DRY_RUN 모드 - 실제 삭제 안 됨')
    Logger.log('실제 삭제하려면 compactEmptyRows 함수의 DRY_RUN = false 로 변경 후 재실행')
    return
  }

  // 실제 삭제 - 뒤에서부터 (앞에서 삭제하면 행 번호가 밀림)
  Logger.log('')
  Logger.log('=== 실제 삭제 시작 ===')
  const lock = LockService.getScriptLock()
  try {
    lock.waitLock(30000)

    // 연속 구간으로 묶어서 삭제 (삭제 횟수 최소화)
    const ranges = []
    let s = emptyRowNumbers[0]
    let e = emptyRowNumbers[0]
    for (let i = 1; i < emptyRowNumbers.length; i++) {
      if (emptyRowNumbers[i] === e + 1) {
        e = emptyRowNumbers[i]
      } else {
        ranges.push({ start: s, end: e })
        s = emptyRowNumbers[i]
        e = emptyRowNumbers[i]
      }
    }
    ranges.push({ start: s, end: e })

    // 뒤에서부터 삭제
    ranges.reverse()
    for (const r of ranges) {
      const count = r.end - r.start + 1
      sheet.deleteRows(r.start, count)
      Logger.log('삭제 완료: ' + r.start + '행부터 ' + count + '행')
    }

    Logger.log('')
    Logger.log('✅ 모든 빈 행 삭제 완료')
    Logger.log('새 getLastRow(): ' + sheet.getLastRow())
  } catch (err) {
    Logger.log('❌ 삭제 실패: ' + err.message)
    throw err
  } finally {
    try { lock.releaseLock() } catch (e) {}
  }
}
