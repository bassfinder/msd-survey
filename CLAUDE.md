# 근골격계 증상조사 · 유해요인조사 정기보고서 앱

수달코리아 EHS. 근골격계부담작업 유해요인조사(안전보건규칙 제657조, 3년 주기 정기 + 수시)를
설문 수집부터 결과보고서(Excel/PDF)까지 한 곳에서 처리한다. soudal-internal-app 스킬 레시피(Flask+SQLite, 무과금 번역, KST, Railway) 기반.

## 실행
```
py app.py            # http://localhost:5003  (관리자 /admin, 기본 암호 msd2026 → 환경변수 ADMIN_PASSWORD)
```
- DB: 앱 폴더 `msd.db`(SQLite). `SQLITE_PATH`로 다른 파일 지정 가능, 배포 시 `DATABASE_URL`(PostgreSQL).
- 환경변수: `ADMIN_PASSWORD`, `SECRET_KEY`, `COMPANY_NAME`(기본 수달코리아(주)), `SITE_NAME`(기본 안성공장).
- **배포(GitHub/Railway)는 사용자 명시 승인 후에만.** 준비 완료 — 절차는 `DEPLOY.md` (수달세이프 프로젝트 안 별도 서비스 + 전용 Postgres, 링크만 연결)
- 클라우드 판정(_on_cloud) = RAILWAY_* 환경변수 또는 postgres DATABASE_URL. 클라우드에서 ADMIN_PASSWORD 없으면 로그인 차단. `SAFETY_SURVEY_URL` 설정 시 완료 화면에 안전설문 링크.
- 로컬 → 서버 데이터 이전: `copy_db.py --only-master` (작업·단위작업·번역·직원명단). PDF 폰트는 static/fonts/NanumGothicBold.ttf(OFL) 동봉.

## 파일
| 파일 | 역할 |
|---|---|
| `survey_def.py` | 증상조사표 문항·선택지 코드, 한국어 원문(KO), **판정 기준 `judge_part()`**, 부담작업 11호, 유해요인조사표 척도·유해요인·평가도구 |
| `app.py` | 모델(Round/Job/Assessment/Photo/Improvement/Response), 설문·관리자 라우트, `analyze()` 공통 집계 |
| `reports.py` | `build_excel` / `build_pdf` / `build_raw_excel` — 모두 `analyze()` 결과 하나만 입력. 별지 제1호 내용은 `f1_sections()`가 Excel·PDF 공용으로 만듦 |
| `make_i18n.py` | KO → 5개 언어 번역 사전 생성 → `i18n/survey_texts.json` (사람이 직접 고쳐도 됨, 관리자 `/admin/reload-texts`로 반영) |
| `templates/` | survey*(직원용 모바일), admin_*(관리자) |

## 업무 흐름
1. 조사회차 생성 → 2. 작업 마스터(부서/공정/작업/인원) 등록 = 설문 선택지 → 3. 회차 "접수 시작" + QR 출력
4. 직원 설문(6개 언어: ko/en/ru/uz/ne/km, 자유기재는 한국어 자동번역 병행저장, 건강정보 수집 동의 필수)
5. 작업별 유해요인조사표 [별지 제1호서식]: 가.조사개요 / 나.작업장상황(변경없음·변경있음) / 다.작업조건(1단계 단위작업, 2단계 단위작업별 부담작업(호)+작업강도A×빈도B, 3단계 원인분석, 3-1 인간공학평가(선택)) / 라.결과요약. 부담작업 체크리스트는 2단계에 통합(Assessment.checklist는 저장 시 tasks에서 자동 생성)
6. 대시보드: 부위별/부서별/작업별, 개선 우선순위 자동 제안(관리대상→NIOSH→A×B→부담작업 수)
7. 개선계획서(우선순위 작업에서 [개선계획 +] → 문제점·근로자 의견 자동 채움)
8. 정기보고서: 표지 → 결과 요약 → 1 조사개요 → 2 체크리스트 → 3 유해요인조사표 → 4 증상조사결과 → 5 종합판정 → 6 개선계획서 → 7 종합의견 + 원자료

## 판정 기준 (survey_def.judge_part) — KOSHA H-30-2008, 2023년 위탁조사 결과표와 동일
- 근거: KOSHA GUIDE E-G-1-2025 <표2> 근골격계질환 증상구분
- 통증호소자('pain'): 기간 ≥1주일 AND 빈도 ≥월1회 AND 정도 심한/매우 심한
- 관리대상자('manage'): (기간 ≥1주일 OR 빈도 ≥월1회) AND 정도 ≥중간, 통증호소자 제외
- 그 외 통증 응답('complaint')은 결과표에서 '정상'. 결과표 표기는 S.KOSHA_KO
- **서식 기준 = 고시 원본 HWP [별지 제1·2호서식](2018.2.9 신설)** 문구 그대로. 증상조사표 통증표 6문항(지난 1주일 포함), 유해요인조사표 작업부하(A)·업무변화·줄음/늘어남/기타. KOSHA 2025 가이드판 아님.
- **절차 기준 = 노동부 근골격계질환 예방업무 편람(2010)**: 체크리스트(단위작업·근로자수·○/×) → 유해요인조사표 → 증상조사 → 유해도 평가·우선순위 → 개선계획(사업주·근로자대표 확인) → 기록보존
- 보고서에 '4-1.증상조사결과표' (2023 결과표와 같은 열: 부위별 정상/관리대상자/통증호소자 + 전체 결과)

## 종합판정 (analyze() 안)
- 부위 부합성: S.BURDEN_PARTS(부담작업 호 → 부위) vs 작업별 관리대상자·통증호소자 부위 → match / unexpected
- 2×2: 유해요인 높음 = 부담작업 AND A×B ≥ S.EXPOSURE_SCORE_HIGH, 증상 많음 = 통증호소자≥1 OR 관리대상 비율 ≥ 사업장 평균 → S.QUADRANT 1~4, 5(조사표 미작성), 0(응답 없음)
- 개인요인: A['personal'] (과거상해 부위 일치·질병·취미 HOBBY_PARTS·가사노동)
- 보고서: 요약 1쪽(summary_data) + 5장 종합판정(judge_rows, JUDGE_BASIS). PDF 기호는 ◎△※ (✔⚠ 글꼴 없음)

## 모바일 작업 선택 (단위작업 프리셋)
- Job.units = 단위작업 프리셋(작업명·위치·작업내용·기본 부담작업 호), Job.units_tr = 언어별 번역 캐시
- 설문 III에서 직원이 단위작업 체크 + 작업부하(A)·작업빈도(B) → Response.units. `unit_stats(round_id, job_id)`가 평균 계산 → 조사표 2단계 [평균 적용]
- 엑셀 '작업별 부하·빈도 응답' 시트 = 응답자×작업 long format (분석용)
- make_i18n.py 끝에서 translate_job_units()로 단위작업 번역도 채움
- 작업자 자가 체크: 단위작업마다 하루시간·무게·횟수·자세·가장 힘든 점 → S.suggest_items()로 부담작업 호 후보 → unit_stats()가 후보 집계(sug, 과반 sug_major)·힘든 점(hz, txt) → 조사표 [과반 후보 체크]·[원인분석 초안 넣기]

## 주의
- 유해요인조사표 2단계는 단위작업 카드(.unit) 구조. 체크박스 이름 t_{카드순번}_{호}는 저장 직전 JS renumber()로 다시 매김 — 카드 순서와 t_name 순서가 일치해야 함.
- 모델명 `Response`가 Flask Response와 겹쳐서 Flask 쪽은 `HttpResponse`로 import.
- Jinja에서 dict의 `items` 키는 `x.get('items')`로 (x.items는 dict 메서드).
- 구글 웹번역은 병렬·연속 호출 시 PC IP를 몇 시간 차단 → make_i18n은 묶음 전송, 차단 감지 시 즉시 중단.
- 번역 사전에 없는 문구는 한국어로 표시(폴백). 현재 5개 언어 수동 번역 완료(i18n/manual_*.py 원본). survey_texts.json 을 고치면 앱이 자동으로 다시 읽음.
- 설문 성명 = 직원명단(Employee) 선택 → Response.employee_id. 사번 칸 없음.
- 보고서 결재란: 담당/검토/승인 (reports.py 표지, Excel·PDF 두 곳).
