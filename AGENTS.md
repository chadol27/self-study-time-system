# 저장소 작업 지침

## 먼저 읽을 문서

- 구현 전에 `PLAN.md`를 읽고 출결 상태, 시트 구조, 인증, 동시성, 날짜 경계 규칙을 따른다.
- 설치·배포·운영 절차는 `README.md`를 따른다. 문서와 실행 코드가 다르면 `src/appsscript.json`, `.claspignore`, 실제 코드를 우선 확인한다.
- 코드를 수정하면 변경된 동작과 정책을 `PLAN.md`에, 설치·배포·운영 영향을 `README.md`에 함께 반영한다.

## 작업 환경

- 이 저장소는 컨테이너 바운드 Google Apps Script 웹 앱이다. 의존성은 `pnpm install`로 설치한다.
- Apps Script에 push되는 manifest, `.gs`, `.html` 파일은 모두 `src/`에 두며 `.clasp.json`의 `rootDir`을 `src`로 유지한다.
- 검사는 `pnpm check`로 실행하며 순서는 `typecheck` → `lint:gs` → `lint:html`이다. 자동 테스트는 없고 실제 Apps Script 템플릿 평가와 런타임은 수동 검증 대상이다.
- `.clasp.json`은 로컬 전용이며 Script ID를 노출하거나 커밋하지 않는다. 배포 반영 명령은 검사 후 `clasp push`를 수행하는 `pnpm deploy`다.
- 날짜 계산은 시스템 로컬 시간 대신 `APP.TZ`와 `DateUtils.gs`를 사용하며 Apps Script와 스프레드시트 시간대를 `Asia/Seoul`로 유지한다.

## 코드 변경 원칙

- 공개 `google.script.run` 함수는 `publicCall_()`의 `{ ok, data/error }` 형식을 유지하고, 클라이언트 입력과 권한은 서버에서 다시 검증한다.
- 시트 쓰기는 `LockService`로 보호하고 잠금 후 최신 상태를 다시 읽는다. 기존 데이터나 알 수 없는 열을 자동 삭제·덮어쓰지 않는다.
- 화면과 사용자 메시지는 한국어로 유지하고 비밀번호·세션 토큰을 URL, 오류, 기록 시트에 남기지 않는다.

## 검증

- 스프레드시트의 `스공시 관리 → 시트 초기화 및 검사`를 실행한 뒤 실제 웹 앱에서 변경한 흐름을 점검한다.
- 날짜·출결 변경은 `PLAN.md`의 `핵심 테스트 항목`을 기준으로 검증한다.
