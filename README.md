# FTP Lite - FANUC 장비 FTP 파일 전송 (간단 버전)

설비의 FTP 데이터 서버와 내 PC 폴더 사이로 파일을 올리고 내려받는 창 하나만 있는 버전입니다.
모니터링(FOCAS), 시간 알람, 장비 패널 정보, 휴대폰 연결, 설정 백업, 사용 기간(라이선스)은 없습니다.

저장소 맨 위(루트)에 아래 파일이 있어야 합니다.

- `FTP_Lite.pyw`
- `FTP_Lite.ico`
- `.github/workflows/build-exe.yml`

## 빌드
파일을 올리면(push) 자동으로 빌드됩니다. Actions 탭 → 맨 위 실행 → 아래 Artifacts 의 `FTP_Lite-exe` 를 받아 압축을 풀면 EXE 입니다.
수동 실행: Actions → Build EXE → Run workflow.

- `FTP_Lite.exe` : 일반 (윈도우 임시 폴더에 풀림)
- `FTP_Lite_tmp.exe` : 임시 폴더가 막힌 PC 용 (C:\ProgramData 에 풀림)

## 설정 파일
`%LOCALAPPDATA%\FTP Lite\Settings.json` (기존 FTP Manager 와 폴더가 달라서 서로 설정을 건드리지 않습니다).
