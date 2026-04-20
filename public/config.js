/**
 * 미담사진관 태블릿 웹 - 설정 파일
 *
 * 2호점 복사 시 이 파일만 수정:
 * 1. APPS_SCRIPT_URL - 새로 배포한 Apps Script Web App URL
 * 2. API_TOKEN - Apps Script CONFIG.API_TOKEN과 일치시킴
 * 3. SHOP_NAME - 매장명 (선택)
 */

window.APP_CONFIG = {
  // Apps Script Web App 배포 URL
  // 배포 -> 새 배포 -> 웹 앱 생성 후 받은 URL
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbz7hsAG93IqpmCtpSekgdCxCPYz35BFNerj93nVDe_rOKoPT-OVm67JVkLaJTNLIIPRlA/exec',

  // Apps Script CONFIG.API_TOKEN과 반드시 동일 문자열
  API_TOKEN: 'midam-2026-secret-token',

  // 매장명
  SHOP_NAME: '미담사진관',

  // 디버그 모드 - border 표시 여부 (개발 중 true / 배포 시 false)
  DEBUG_BORDER: false,

  // 대기리스트 자동 새로고침 주기 (ms) - 0이면 비활성화
  AUTO_REFRESH_MS: 30000,

  // 요청 타임아웃 (ms)
  REQUEST_TIMEOUT_MS: 15000
}
