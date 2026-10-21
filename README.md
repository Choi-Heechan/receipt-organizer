# receipt-organizer

영수증(PDF·이미지)을 사람별 폴더 단위로 읽어 엑셀·JSON·대시보드를 만드는 Claude Code 플러그인입니다.
領収書（PDF・画像）を個人フォルダ別に読み取り、Excel・JSON・ダッシュボードを作成する Claude Code プラグインです。

- [한국어](#한국어)
- [日本語](#日本語)

---

## 한국어

### 무엇을 하나요

상위 폴더 아래 사람별 폴더에 넣은 영수증(`.jpg .jpeg .png .pdf`)을 Claude가 읽어서, 상위 폴더 바로 아래에 다음 파일을 만듭니다(YYYYMMDD = 실행일).

| 파일 | 내용 |
|---|---|
| `receipt_<폴더명>_YYYYMMDD.xlsx` | 사람별 영수증 목록(날짜·상호·등록번호·금액·세율·결제수단) |
| `list_receipt_YYYYMMDD.xlsx` | 폴더별 건수·합계 금액 요약 |
| `receipt_YYYYMMDD.json` | 전체 결과 데이터 |
| `dashboard_YYYYMMDD.html` | 브라우저로 보는 사람별 카드 화면 |

읽지 못한 칸은 `読み取りエラー`(빨간색), 읽었지만 확인이 필요한 칸은 `要確認`(노란색)으로 표시됩니다. 같은 날 다시 실행하면 덮어쓰지 않고 `_2`, `_3`…이 붙습니다.

### 설치

Claude Code에서 다음을 입력합니다(GitHub 계정 불필요).

```
/plugin marketplace add Choi-Heechan/receipt-organizer
/plugin install receipt-organizer@receipt-tools
```

> 같은 PC에 예전 `receipt-processor` 스킬이 있으면 호출 문구가 겹칩니다. 그 스킬은 비활성화(또는 삭제)해 주세요.

### 폴더 구조

```
영수증_2026년10월/          ← 실행할 때 이 경로를 지정
├── 홍길동/
│   ├── 20261003_택시.jpg
│   └── 20261005_식당.pdf
└── 김철수/
    └── ...
```

- 사람별 폴더 안의 더 깊은 폴더도 그 사람의 영수증으로 처리합니다.
- 하위 폴더 없이 파일만 있는 폴더를 지정하면 그 폴더를 한 사람으로 처리합니다.
- 상위 폴더에 직접 놓인 파일, 빈 폴더, HEIC 등 지원하지 않는 파일은 처리하지 않고 완료 보고에 알려 줍니다.
- 영수증 1파일 = 영수증 1장으로 가정합니다.

### 사용법

Claude Code에서 상위 폴더 경로와 함께 요청합니다.

```
영수증 정리해줘 C:\Users\...\領収書_2026年10月
```

또는 상위 폴더로 이동한 뒤(터미널에서 `cd` 후 `claude` 실행, 데스크톱 앱에서는 세션 폴더로 지정) 경로 없이 요청해도 됩니다. 이 경우 Claude가 현재 폴더에서 찾은 개인 폴더·영수증 수를 보여 주고 진행 여부를 한 번 확인합니다.

```
영수증 처리해줘
```

끝나면 폴더별 건수·합계 금액 표와 생성 파일 경로를 알려 줍니다. `dashboard_YYYYMMDD.html`을 더블클릭하면 브라우저에서 열립니다(인터넷 불필요).

### Python 준비

파일 생성에 Python 3(추가 패키지 불필요)를 사용합니다. 처음 실행할 때 Claude가 자동으로 확인하고, 없으면 설치를 안내합니다.

- Windows: `winget install Python.Python.3.12` (동의하면 Claude가 실행합니다). winget을 쓸 수 없으면 https://www.python.org/downloads/ 에서 설치하고, 첫 화면의 「Add python.exe to PATH」에 체크하세요. 설치 후에는 Claude를 다시 시작해야 인식될 수 있습니다.
- Mac: 터미널에서 `xcode-select --install` 을 실행하거나 macOS가 띄우는 개발자 도구 설치 창을 따르세요.

### 개인정보

결과 파일(엑셀·JSON·대시보드)은 모두 PC의 지정 폴더에만 저장되며, 이 플러그인은 별도 서버나 외부 OCR 서비스로 데이터를 보내지 않습니다. 영수증 판독은 평소 Claude와 대화할 때와 같은 방식으로 Claude가 직접 합니다(이미지가 Claude에게 전달되는 것은 일반 Claude Code 사용과 같습니다).

### 업데이트

```
/plugin marketplace update receipt-tools
```

그 후 Claude Code를 다시 시작합니다. (`/plugin` 메뉴에서도 업데이트할 수 있습니다.)

### 관리자용 메모

새 버전을 배포할 때마다 `plugins/receipt-organizer/.claude-plugin/plugin.json`의 `"version"`을 올려 주세요(예: `0.1.1` → `0.1.2`). 버전이 그대로면 사용자가 업데이트해도 변경 사항을 받지 못합니다.

---

## 日本語

### できること

上位フォルダの下の個人フォルダに入れた領収書（`.jpg .jpeg .png .pdf`）を Claude が読み取り、上位フォルダ直下に次のファイルを作成します（YYYYMMDD＝実行日）。

| ファイル | 内容 |
|---|---|
| `receipt_<フォルダ名>_YYYYMMDD.xlsx` | 個人別の領収書一覧（日付・店名・登録番号・金額・税率・支払区分） |
| `list_receipt_YYYYMMDD.xlsx` | フォルダ別の件数・合計金額のまとめ |
| `receipt_YYYYMMDD.json` | 全結果データ |
| `dashboard_YYYYMMDD.html` | ブラウザで見る個人別カード画面 |

読み取れなかった項目は `読み取りエラー`（赤）、読み取ったが確認が必要な項目は `要確認`（黄）で表示されます。同じ日に再実行すると上書きせず `_2`、`_3`… を付けます。

### インストール

Claude Code で次を入力します（GitHub アカウント不要）。

```
/plugin marketplace add Choi-Heechan/receipt-organizer
/plugin install receipt-organizer@receipt-tools
```

> 同じ PC に以前の `receipt-processor` スキルがある場合、呼び出しの言葉が重なります。そのスキルは無効化（または削除）してください。

### フォルダ構成

```
領収書_2026年10月/          ← 実行時にこのパスを指定
├── 山田太郎/
│   ├── 20261003_タクシー.jpg
│   └── 20261005_会食.pdf
└── 佐藤花子/
    └── ...
```

- 個人フォルダの中のさらに深いフォルダも、その人の領収書として処理します。
- サブフォルダがなくファイルだけのフォルダを指定した場合、そのフォルダを 1 人分として処理します。
- 上位フォルダに直接置かれたファイル、空のフォルダ、HEIC など未対応のファイルは処理せず、完了報告でお知らせします。
- 1 ファイル＝領収書 1 枚として扱います。

### 使い方

Claude Code で上位フォルダのパスと一緒に依頼します。

```
領収書をまとめて C:\Users\...\領収書_2026年10月
```

または上位フォルダに移動してから（ターミナルで `cd` してから `claude` を起動、デスクトップアプリではセッションのフォルダに指定）、パスなしで依頼することもできます。この場合 Claude が現在のフォルダで見つけた個人フォルダ・領収書の数を表示し、処理してよいか一度確認します。

```
領収書処理して
```

完了するとフォルダ別の件数・合計金額の表と、作成したファイルのパスを表示します。`dashboard_YYYYMMDD.html` をダブルクリックするとブラウザで開きます（インターネット不要）。

### Python の準備

ファイル作成に Python 3（追加パッケージ不要）を使います。初回実行時に Claude が自動で確認し、見つからなければインストールを案内します。

- Windows：`winget install Python.Python.3.12`（同意すると Claude が実行します）。winget が使えない場合は https://www.python.org/downloads/ からインストールし、最初の画面で「Add python.exe to PATH」にチェックを入れてください。インストール後、Claude の再起動が必要な場合があります。
- Mac：ターミナルで `xcode-select --install` を実行するか、macOS が表示する開発者ツールのインストール画面に従ってください。

### 個人情報

結果ファイル（Excel・JSON・ダッシュボード）はすべて PC の指定フォルダにのみ保存され、このプラグインが外部サーバーや外部 OCR サービスにデータを送信することはありません。領収書の読み取りは、普段 Claude と会話するときと同じ方法で Claude 自身が行います（画像が Claude に渡るのは通常の Claude Code の利用と同じです）。

### アップデート

```
/plugin marketplace update receipt-tools
```

その後 Claude Code を再起動してください。（`/plugin` メニューからも更新できます。）
