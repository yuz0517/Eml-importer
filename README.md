# EML → Outlook PST 일괄 가져오기

그룹웨어에서 내보낸 `.eml` 파일을 **Outlook에 일괄 삽입**하는 Windows 도구입니다.  
**Python 설치 불필요** — `.exe` 파일 하나만 다운로드하면 바로 실행됩니다.

---

## ⬇️ 다운로드

1. 이 저장소의 **[Releases](../../releases)** 탭 클릭
2. 최신 버전의 `EML_Importer.exe` 다운로드
3. 더블클릭으로 실행

---

## 🖥️ 사용법

1. **Outlook을 먼저 실행**하세요
2. `EML_Importer.exe` 실행
3. **① EML 폴더** 선택 — 그룹웨어에서 내려받은 `.eml` 파일 모음 폴더
4. **② Outlook 폴더** 선택 — 삽입할 대상 폴더 (보낸 편지함 등)
5. **▶ 가져오기 시작** 클릭

> "하위 폴더 포함" 체크 시 하위 폴더 내 EML까지 모두 처리합니다.

---

## ✅ 처리 항목

| 항목 | 지원 |
|------|------|
| 메일 제목 | ✓ |
| 발신자 / 수신자 / 참조 | ✓ |
| 발송 일시 (원본 날짜 보존) | ✓ |
| 본문 HTML / 텍스트 | ✓ |
| 첨부파일 | ✓ |
| 하위 폴더 재귀 처리 | ✓ |
| 실시간 진행률 + 로그 | ✓ |

---

## ⚠️ 요구사항

- Windows 10 / 11
- Microsoft Outlook 설치 (365 / 2019 / 2021)
- Python **불필요**

---

## 🔧 직접 빌드하려면

```bat
pip install pywin32 pyinstaller
pyinstaller --onefile --windowed --name EML_Importer eml_importer.py
```

또는 이 저장소를 Fork 후 GitHub Actions가 자동으로 빌드합니다.
