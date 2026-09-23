# 배포 절차 (준비 완료 · 아직 미배포)

> 배포(GitHub push, Railway)는 **사용자 승인 후에만** 진행한다.

## 구성 (권장)
- **수달세이프와 같은 Railway 프로젝트 안에, 별도 서비스로** 추가
  - 서비스: `msd-survey` (이 폴더) + **전용 PostgreSQL** (수달세이프 DB와 분리)
  - 이유: 증상조사 = 실명 건강정보 / 수달세이프 안전설문 = 무기명 → DB·권한을 섞지 않는다
- 수달세이프에는 **링크만** 추가 (홈 메뉴 또는 안전설문 완료 화면 → 근골격계 증상조사)

## 1. GitHub
1. 새 비공개 저장소 생성 (예: `bassfinder/msd-survey`)
2. 이 폴더를 push — `msd.db`, `__pycache__` 는 `.gitignore`로 제외됨

## 2. Railway
1. 수달세이프 프로젝트 → New → GitHub Repo → `msd-survey`
2. New → Database → PostgreSQL (이 서비스 전용) → `DATABASE_URL` 연결
3. Variables

| 변수 | 값 | 비고 |
|---|---|---|
| `ADMIN_PASSWORD` | 새 관리자 암호 | **필수** — 없으면 서버에서 로그인 차단 |
| `SECRET_KEY` | 임의의 긴 문자열 | 세션 암호화 |
| `DATABASE_URL` | (Postgres 연결 시 자동) | |
| `SURVEY_OPEN` | `0` (테스트 기간) → 직원 공개 시 `1` | 0이면 직원은 '조사 없음', 관리자만 로그인해서 전 기능 테스트 |
| `SAFETY_SURVEY_URL` | 수달세이프 안전설문 주소 | 넣으면 증상조사 완료 화면에 '안전 설문도 참여하기' 버튼 |
| `COMPANY_NAME` / `SITE_NAME` | 수달코리아(주) / 안성공장 | 기본값과 같으면 생략 |

4. Settings → Networking → Generate Domain → `https://msd-survey-xxxx.up.railway.app`
5. `https://…/healthz` 가 `{"ok": true}` 이면 정상

## 3. 기본 데이터 옮기기 (작업 12개·단위작업·번역·직원명단)
- **권장(간단)**: PC 앱 [작업·유해요인조사표] → ⬇ 내보내기(JSON) → 서버 관리자 화면 같은 곳에서 ⬆ 가져오기
  - 미리 만들어 둔 파일: `05_근골격계\근골격계_기본데이터_서버업로드용.json` (직원 이름 포함 — 저장소에 올리지 말 것)
- 대안: `set DATABASE_URL=<Postgres Public URL>` 후 `py copy_db.py --only-master`

## 4. 운영 시작
1. 서버 관리자 화면 → [조사회차] 새로 만들기 → [접수 시작] (SURVEY_OPEN=0 동안은 관리자만 보임)
1-1. 테스트 끝나면 테스트 응답 삭제 → Railway Variables `SURVEY_OPEN=1` → 직원 공개
2. [QR] → 새 주소로 QR 출력 (로컬 QR은 사용 불가)
3. 수달세이프에 링크 추가 (아래)

## 5. 수달세이프 쪽 연결 (별도 작업, 승인 후)
- 안전설문 완료 화면 또는 홈 메뉴에 버튼 1개:
  `<a href="https://msd-survey-xxxx.up.railway.app/survey">근골격계 증상조사 하기</a>`
- 이름·응답 등 어떤 정보도 넘기지 않는다(무기명 유지).
- 수달세이프 코드는 GitHub `bassfinder/soudal-safe` 가 정본(로컬 폴더에는 설문 코드 없음) — 먼저 clone 후 수정.

## 점검 결과 (로컬, 2026-09-23)
- PDF 한글: 서버엔 맑은고딕이 없어 `static/fonts/NanumGothicBold.ttf`(OFL) 동봉 → 자동 사용
- 클라우드 기본 암호 차단, `/healthz`, copy_db (SQLite→SQLite 검증) 통과
- 증상조사 자유기재 번역은 서버에서 deep_translator(무과금) 호출
