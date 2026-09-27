# 메일 배치 — 엑셀 명단 개인별 메일 일괄발송

> **명단 속 수십 명에게 이름·첨부만 바꿔 메일을 한 통씩 보내던 일을 한 번에 끝냅니다.**

엑셀 명단으로 한 사람 한 사람에게 **개인화된 메일**을 자동 발송하는 Windows 데스크톱 앱.
수료증·안내문·공문 발송에 딱 맞습니다. API 키·서버 불필요, 완전 로컬 동작. **무료**입니다.

## 사용 모습

<!-- 사용 GIF 준비 중 -->

## 내려받기

- **⬇ 다운로드: [mailbatch.vercel.app](https://mailbatch.vercel.app)** · [Releases](https://github.com/stanlee7/mail-batch/releases/latest)

> [HWP 배치](https://hwp-batch.vercel.app)로 수료증 PDF를 대량 생성했다면,
> 이 도구로 한 사람씩 메일 발송까지 끝낼 수 있습니다.

## 기능
- **개인화 발송**: 엑셀 명단 + `{{컬럼명}}` 자리표시 제목/본문 → 행마다 치환해 개별 발송
- **개인별 첨부**: `{이름}_수료증.pdf` 같은 파일명 패턴으로 폴더에서 사람별 첨부 자동 매칭 (공통 첨부도 가능)
- **안전장치**: 테스트 발송(1행→내 주소), 발송 전 미리보기, 주소 오류 검출, 발송 간격 조절, 실패만 재발송, 결과 엑셀 저장
- **계정 프리셋**: 네이버 / 지메일 / 다음(카카오) / 직접 입력 — 앱 비밀번호 안내 포함
- 비밀번호는 저장을 켠 경우에만 Windows DPAPI로 암호화해 이 PC에만 보관

## 사용법
1. **명단 탭** — 엑셀 선택 (첫 행 = 컬럼명, 예: 이름/이메일/과정명). 이메일 컬럼은 자동 감지
2. **메일 작성 탭** — 제목/본문 입력, 컬럼 버튼 클릭으로 `{{이름}}` 삽입. 첨부 설정 후 미리보기로 확인
3. **발송 탭** — 메일 서비스 선택 + 아이디/앱 비밀번호 입력 → **테스트 발송**으로 확인 → 전체 발송

### 앱 비밀번호?
네이버·지메일은 보안 정책상 일반 비밀번호로 SMTP 로그인이 안 됩니다.
- 네이버: 메일 → 환경설정 → POP3/SMTP 사용 켜기 (2단계 인증 시 앱 비밀번호 발급)
- 지메일: Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호

⚠ 무료 메일 계정은 일일 발송 한도(약 500건)가 있습니다. 대량 발송 시 간격을 2초 이상으로 두세요.

## 개발 실행
```
pip install openpyxl PySide6
python app.py
```

## 구조
- `mailbatch/engine.py` — 코어 엔진 (엑셀 읽기, 템플릿 렌더, 첨부 매칭, SMTP 발송)
- `mailbatch/gui.py` — PySide6 GUI (발송은 QThread 워커)
- `mailbatch/providers.py` — SMTP 프리셋 / `mailbatch/settings.py` — 설정·DPAPI 암호화
- `tests/test_engine.py` — 오프라인 단위 테스트
- `tests/test_send_local.py` — 로컬 SMTP 서버 통합 테스트 (`pip install aiosmtpd`)
- `tests/smoke_gui.py` — GUI 스모크 (offscreen)

## 빌드
```
python -m PyInstaller MailBatch.spec --noconfirm
dist\MailBatch\MailBatch.exe --selftest   # SELFTEST: SUCCESS 확인
```

## 우리 팀에 도입·교육이 필요하면

에이전엘은 기업·기관 AI 교육과 AX 업무 도구 구축을 합니다. 이 도구를 팀 업무에 맞게 적용하거나 사용 교육이 필요하면 → [에이전엘 견적 요청](https://agenaile.com/?utm_source=github&utm_medium=readme&utm_campaign=tools&utm_content=mail-batch#quote)
