# 스공시 좌석 및 출결 관리 시스템

Google Spreadsheet에 연결해 사용하는 컨테이너 바운드 Google Apps Script 웹 앱입니다. 상세 정책은 `AGENTS.md`를 참고하세요.

## 설치

1. 대상 스프레드시트에서 **확장 프로그램 → Apps Script**를 열어 프로젝트를 만듭니다.
2. 저장소 루트에서 `pnpm install`을 실행한 뒤 `pnpm clasp login`으로 Google 계정에 로그인합니다.
3. `.clasp.json.example`을 `.clasp.json`으로 복사하고 Apps Script의 스크립트 ID를 입력합니다. `rootDir`은 배포 파일이 모인 `src`로 유지합니다.
4. `pnpm deploy`로 정적 검사를 통과한 파일을 올립니다.
5. Apps Script의 **프로젝트 설정 → 스크립트 속성**에 아래 값을 추가합니다.

   - `TEACHER_PASSWORD`: 교사 비밀번호
   - `TOTAL_SEATS`: 예: `20`
   - `PERIOD_1_START_TIME`: 예: `18:00`
   - `PERIOD_2_START_TIME`: 예: `19:20`
   - `PERIOD_3_START_TIME`: 예: `20:50`

6. 스프레드시트를 새로 열고 **스공시 관리 → 시트 초기화 및 검사**를 실행합니다.
7. **스공시 관리 → 일일 트리거 설치**를 한 번 실행하고 권한을 승인합니다.
8. Apps Script에서 **배포 → 새 배포 → 웹 앱**을 선택합니다. 실행 사용자는 배포자, 액세스 권한은 링크를 아는 누구나로 설정합니다.

## 개발 명령

- `pnpm typecheck`: `src/*.gs`를 임시 JavaScript 작업공간에서 TypeScript와 Apps Script 타입으로 검사합니다.
- `pnpm lint:gs`: 모든 `src/*.gs` 파일을 ESLint로 검사합니다.
- `pnpm lint:html`: Apps Script 템플릿 표현식을 전처리한 뒤 HTML 구조와 페이지별 인라인 JavaScript를 검사합니다.
- `pnpm lint`: `lint:gs` 후 `lint:html`을 순서대로 실행합니다.
- `pnpm check`: `typecheck` 후 `lint`를 순서대로 실행합니다.
- `pnpm deploy`: `check`가 성공하면 로컬 `clasp`로 Apps Script 프로젝트에 `push`합니다.

HTML 구조와 인라인 JavaScript는 정적 검사하지만 Apps Script 템플릿의 실제 평가 결과와 런타임 동작까지 검증하지는 않습니다. 배포 후 스프레드시트 메뉴 검사와 실제 웹 앱 흐름을 함께 점검하세요.

Apps Script에 업로드되는 `appsscript.json`, `.gs`, `.html` 파일은 모두 `src/`에 있습니다. 저장소 루트의 문서, 개발 설정, 스크립트는 `.claspignore`와 `rootDir` 설정에 따라 업로드되지 않습니다.

프로젝트와 스프레드시트의 시간대는 모두 `Asia/Seoul`로 맞춰야 합니다. 명부의 학번·이름을 바꿔야 할 때는 기존 행 E:P를 모두 비운 뒤 새 행을 추가합니다. 미운영일은 `설정` 시트 A열에 실제 날짜값으로 입력합니다. 일일 트리거는 오늘부터 30일 후까지 운영일의 출결 열을 생성하고, 같은 범위에 이미 생성된 미운영일 열은 삭제합니다. 과거 출결 열은 삭제하지 않습니다.

## 명부 입력

`명부` 시트 3행부터 학번, 이름, 좌석번호, 전화번호 뒤 4자리와 월~목 1~3교시 신청값을 입력합니다. 신청은 `1`, 미신청은 `0` 또는 빈 값입니다. Q열 내부 키는 시스템이 자동 생성하고 숨깁니다.

배포 후에는 학생 로그인, 사전 결석 등록/취소, 교사 로그인, 좌석 조회, 결석 처리/복원을 실제 계정 흐름으로 점검하세요. 날짜 헤더에는 연도가 포함되므로 새해에 기존 출결 열을 지울 필요가 없습니다. 보관이 필요하면 스프레드시트 전체를 연도별로 복사합니다.
