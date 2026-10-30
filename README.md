# 국회 유통 관련 법안 모니터링 (자동 실행)

평일 아침 7시 30분경(한국시간), GitHub 서버가 열린국회정보 OpenAPI에서 제22대 국회 법률안과 진행중 입법예고를 받아
관심 키워드 법안을 추리고, 직전 실행 대비 **신규 발의·진행상황 변동**을 보고서로 만들어 웹 주소에 게시합니다.
휴대폰에서는 그 주소만 열면 됩니다.

---

## 최초 설정 (PC에서 약 15분, 한 번만)

### 1. GitHub 가입
- https://github.com 에서 무료 계정을 만듭니다. (이미 있으면 생략)

### 2. 저장소 만들기
- 오른쪽 위 **+** → **New repository**
- Repository name: 예) `bill-monitor` (회사명이 드러나지 않는 이름 권장)
- **Public** 선택 → **Create repository**

> 무료 계정에서 웹 게시(GitHub Pages)를 쓰려면 Public이어야 합니다. 저장소 내용(프로그램·키워드 목록·보고서)은
> 누구나 볼 수 있지만, **인증키는 3단계의 Secret에 넣으므로 공개되지 않습니다.**
> 보고서 내용은 모두 국회 공개 정보입니다.

### 3. 인증키 등록 (비공개 보관)
- 저장소 화면 → **Settings** → 왼쪽 **Secrets and variables** → **Actions**
- **New repository secret**
  - Name: `ASSEMBLY_API_KEY`
  - Secret: 열린국회정보에서 받은 인증키
- **Add secret**

### 4. 파일 올리기
- 저장소 첫 화면 → **Add file** → **Upload files** (빈 저장소면 "uploading an existing file" 링크)
- 압축을 푼 폴더 안의 **모든 항목**(`.github` 폴더, `monitor.py`, `config.json`, `README.md`)을 끌어다 놓기
- 아래 **Commit changes**

> `.github` 폴더가 안 올라가면: **Add file → Create new file** → 파일 이름 칸에
> `.github/workflows/monitor.yml` 입력 → 메모장으로 연 `monitor.yml` 내용을 붙여넣고 Commit

### 5. 첫 실행 (수동)
- 저장소 위쪽 **Actions** 탭 → (안내가 나오면 "I understand… enable" 클릭)
- 왼쪽 **국회 법안 모니터링** → 오른쪽 **Run workflow** → **Run workflow**
- 1~3분 뒤 초록색 체크 ✅ 가 뜨면 성공
- 빨간 ✕ 이면 해당 실행을 눌러 **연결 점검(실패 시에만)** 단계의 내용을 확인 (아래 '문제 해결' 참조)

### 6. 웹 게시 켜기
- **Settings** → 왼쪽 **Pages**
- Source: **Deploy from a branch** / Branch: **main**, 폴더 **/docs** → **Save**
- 1~2분 뒤 같은 화면 위쪽에 주소가 표시됩니다: `https://<아이디>.github.io/bill-monitor/`
- 휴대폰에서 이 주소를 열고 **홈 화면에 추가** 해 두면 앱처럼 쓸 수 있습니다.

---

## 평소 사용
- 평일 아침마다 자동 실행 → 같은 주소에 최신 보고서가 올라옵니다. (GitHub 사정으로 수십 분 늦을 수 있음)
- 보고서 위쪽 버튼: **지난 보고서**(최근 90회분 보관) / **엑셀 받기**
- 실행이 실패하면 GitHub 가입 메일로 실패 알림이 옵니다.
- 지금 바로 갱신하고 싶을 때: Actions → Run workflow (휴대폰 브라우저나 GitHub 앱에서도 가능)

## 키워드·설정 바꾸기
- 저장소에서 `config.json` 클릭 → 연필(✏️) 아이콘 → 수정 → **Commit changes**
- `keyword_groups` : `"관심분야": ["키워드", ...]` — 의안명에 포함되면 추적 (띄어쓰기 무시)
- `exclude_keywords` : 의안명에 포함되면 제외
- `watch_committees` : 키워드와 무관하게 전부 추적할 위원회. 예) `["산업통상자원중소벤처기업위원회"]`
- `api_key` 는 비워 두세요. (Secret 값을 사용)
- 실행 시각 변경: `.github/workflows/monitor.yml` 의 `cron` 값 (UTC 기준, 한국시간 −9시간)

## 문제 해결
| 증상 | 확인 |
|---|---|
| `ERROR-290` | 인증키 오류 → 3단계 Secret 이름·값 재확인 |
| `접속 실패` / 시간 초과 | 국회 서버가 해외(GitHub 서버) 접속을 막는 경우일 수 있음 → 이 경우 PC 실행 방식으로 전환 |
| `결과 저장` 단계에서 403 | Settings → Actions → General → Workflow permissions → **Read and write permissions** 선택 후 Save |
| Pages 주소가 404 | 첫 실행 성공 후 6단계를 했는지, `docs` 폴더가 생겼는지 확인 |

자료: 열린국회정보 OpenAPI(국회사무처) — 의안검색(TVBPMBILL11), 국회의원 발의법률안(nzmimeepazxkubdpn), 진행중 입법예고(nknalejkafmvgzmpt)
