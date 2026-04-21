# ScreenCapture

Windows용 트레이 기반 화면 캡처 도구입니다. 단축키로 영역을 캡처하고, 클립보드 복사, 파일 저장, 스티키 이미지 창 고정, OCR 텍스트 처리, ChatGPT 전송, 배경 제거 같은 후처리를 빠르게 실행할 수 있습니다.

## 주요 기능

- 영역 선택 캡처
- 이미지 클립보드 복사
- 파일 저장 및 저장 경로 클립보드 복사
- 스티키 캡처 창 고정
- 마지막 캡처 영역 반복
- OCR 텍스트 추출 및 복사
- 캡처 이미지에서 원본 링크/파일 열기
- 스티키 이미지 편집: 좌우 반전, 상하 반전, 회전, 다른 이름으로 저장
- 부가기능: ChatGPT로 보내기, 로컬 배경 제거
- AI 이미지 편집은 현재 미구현 상태로 비활성화되어 있습니다.

## 기본 단축키

| 단축키 | 기본 동작 |
|---|---|
| `Ctrl+Shift+A` | 영역 캡처 후 이미지를 클립보드에 복사 |
| `Ctrl+Shift+B` | 영역 캡처 후 파일로 저장하고 저장 경로를 클립보드에 복사 |
| `Ctrl+Shift+S` | 영역 캡처 후 파일로 저장하고 스티키 창으로 표시 |

단축키와 각 단축키의 동작은 트레이 아이콘의 `옵션(O)...` 메뉴에서 변경할 수 있습니다.

## 트레이 메뉴

트레이 아이콘을 우클릭하면 다음 기능을 사용할 수 있습니다.

- `화면 캡쳐 도구(C)`: 단축키 A/B/C 동작을 메뉴에서 직접 실행
- `마지막 캡쳐 영역 반복(L)`: 직전에 선택한 캡처 영역을 다시 캡처
- `스티키 모드(S)`: 캡처 결과를 항상 스티키 창으로 띄우는 모드
- `모든 스티키 닫기(A)`: 열려 있는 스티키 창을 모두 닫기
- `옵션(O)...`: 단축키, 저장 위치, 이미지 형식, OCR 등 설정 변경
- `종료(X)`: 프로그램 종료

## 스티키 창 조작

스티키 창은 캡처 이미지를 화면 위에 고정해서 참고할 때 쓰는 작은 이미지 창입니다.

| 조작 | 동작 |
|---|---|
| 마우스 왼쪽 드래그 | 창 이동 |
| `Shift` + 드래그 | 창 크기 조절 |
| 마우스 휠 | 투명도 조절 |
| `Shift` + 마우스 휠 | 확대/축소 |
| 더블 클릭 | 캡처한 원본 창/링크/파일 열기 또는 원본 위치 클릭 |
| `Esc` 또는 `Delete` | 스티키 창 닫기 |
| 우클릭 | 스티키 창 메뉴 열기 |

## 스티키 창 우클릭 메뉴

- `링크 열기`, `파일 열기`, `위치 열기`: 캡처 원본이 확인된 경우 표시됩니다.
- `캡쳐 경로 복사`, `캡쳐 위치 열기`: 저장된 캡처 파일이 있을 때 표시됩니다.
- `텍스트 복사`, `메모장에 열기`: OCR 또는 클립보드 텍스트가 있을 때 표시됩니다.
- `부가기능`
  - `ChatGPT로 보내기`: 이미지를 클립보드에 넣고 ChatGPT를 열어 붙여넣습니다.
  - `배경 제거`: 로컬 Python `rembg`로 배경을 제거합니다.
  - `AI 이미지 편집 (미구현)`: 추후 구현 예정입니다.
- `이미지 편집`: 좌우 반전, 상하 반전, 회전, 다른 이름으로 저장
- `항상 위(T)`: 스티키 창의 항상 위 상태를 전환
- `닫기`: 현재 스티키 창 닫기

## 설정

설정은 `%APPDATA%\ScreenCapture\settings.json`에 저장됩니다.

기본값:

- 시작 시 Windows와 함께 실행: 켜짐
- 시작 시 최소화: 켜짐
- 캡처 완료음 재생: 켜짐
- OCR 사용: 켜짐
- 저장 폴더: 바탕화면
- 이미지 형식: JPG
- JPG 품질: 95
- 파일명 패턴: `yyyyMMddHHmmss`

## 배경 제거 준비

`부가기능 > 배경 제거`는 로컬 Python 패키지를 사용합니다. 처음 사용할 때 Python과 패키지가 필요합니다.

```powershell
pip install rembg Pillow
```

배경 제거 결과는 저장 폴더에 `yyyyMMddHHmmss_nobg.png` 형식으로 저장되고, 새 스티키 창으로 열립니다.

## 빌드

필요 조건:

- Windows
- .NET SDK

빌드:

```powershell
dotnet build .\ScreenCapture\ScreenCapture.csproj -c Release
```

단일 실행 파일 publish 예시:

```powershell
dotnet publish .\ScreenCapture\ScreenCapture.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

빌드 산출물과 실행 파일은 Git에 커밋하지 않습니다. 배포용 `.exe`는 GitHub Releases 같은 릴리스 첨부 파일로 올리는 방식이 적합합니다.

## 저장소 구조

```text
ScreenCapture/
  ScreenCapture/
    AppSettings.cs
    TrayApp.cs
    StickyWindow.cs
    SettingsForm.cs
    SelectionOverlay.cs
    SourceDetector.cs
    OcrHelper.cs
    AiToolsForm.cs
    ScreenCapture.csproj
```

## 현재 상태

- 기본 캡처, 저장, 클립보드, 스티키 창 기능은 구현되어 있습니다.
- ChatGPT 전송과 로컬 배경 제거는 `부가기능`으로 제공됩니다.
- AI 이미지 편집 API 연동은 비용 문제 때문에 현재 미구현으로 남겨두었습니다.
