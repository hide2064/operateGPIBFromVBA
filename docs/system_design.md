# システム設計書 — operateGPIBFromVBA

## 目次

1. [システム概要](#1-システム概要)
2. [処理ルート](#2-処理ルート)
   - 2.1 [CLI 方式](#21-cli-方式)
   - 2.2 [Flask 方式（通常コマンド）](#22-flask-方式通常コマンド)
   - 2.3 [Flask 方式（MT8821C 専用）](#23-flask-方式mt8821c-専用)
   - 2.4 [LAN 接続の分岐](#24-lan-接続の分岐)
   - 2.5 [リトライ・再接続フロー](#25-リトライ再接続フロー)
   - 2.6 [試験結果の記録フロー](#26-試験結果の記録フロー)
3. [モジュール構成](#3-モジュール構成)
4. [ファイル構成](#4-ファイル構成)
5. [Excel シート仕様](#5-excel-シート仕様)
6. [設定ファイル仕様（settings.ini）](#6-設定ファイル仕様settingsini)
7. [API エンドポイント仕様](#7-api-エンドポイント仕様)
8. [VISA アドレス仕様](#8-visa-アドレス仕様)
9. [プラグイン拡張方式](#9-プラグイン拡張方式)

---

## 1. システム概要

Excel VBA から Python (pyvisa) を経由して GPIB / LAN 機器を制御するシステム。

```
┌─────────────────────────────────────────────────────────┐
│  Excel (VBA)                                            │
│  ┌──────────┐  ┌────────────────┐  ┌─────────────────┐ │
│  │ Config   │  │ Control        │  │ Result          │ │
│  │ シート   │  │ シート         │  │ シート          │ │
│  │ 機器設定 │  │ コマンド実行   │  │ 試験結果記録    │ │
│  └──────────┘  └────────────────┘  └─────────────────┘ │
└───────────────────────┬─────────────────────────────────┘
                        │ 2方式
           ┌────────────┴────────────┐
           │ CLI方式                 │ Flask方式 (推奨)
           ▼                        ▼
   WScript.Shell.Exec()      MSXML2.XMLHTTP
   (プロセス都度起動)          (HTTP POST)
           │                        │
           ▼                        ▼
   gpib_controller.py        server.py (常駐)
           │                        │
           └────────────┬───────────┘
                        ▼
                 pyvisa (VISA ドライバ)
                        │
           ┌────────────┴────────────┐
           ▼                        ▼
      GPIB ボード               LAN (TCP/IP)
           │                        │
           └────────────┬───────────┘
                        ▼
                   計測機器
         (DMM / 電源 / MT8821C 等)
```

### 2 方式の比較

| 項目 | CLI 方式 | Flask 方式 (推奨) |
|------|----------|------------------|
| VBA モジュール | `GpibControl.bas` | `GpibControlHttp.bas` |
| Python 起動 | コマンド毎に起動 | サーバーとして常駐 |
| GPIB 接続 | 毎回 open/close | 接続プールで再利用 |
| リトライ | なし | 失敗時に自動再接続 |
| ログ出力 | なし | ファイル + コンソール |
| 用途 | 試験・デバッグ向け | 本番運用向け |

---

## 2. 処理ルート

### 2.1 CLI 方式

```mermaid
flowchart TD
    A([Excel ボタン押下]) --> B[GpibControl.bas\nExecuteSelectedCommand]
    B --> C[Config シートから\nアドレス・タイムアウト取得]
    C --> D{D/E 列に\nProtocol/Host\nあり?}
    D -- あり --> E[BuildVisaAddress\nでアドレス生成]
    D -- なし --> F[B列のアドレスを使用]
    E --> G
    F --> G[WScript.Shell.Exec\nで Python 起動]
    G --> H[gpib_controller.py\n--address --command]
    H --> I[GenericInstrument.execute]
    I --> J{コマンドが\n'?' で終わる?}
    J -- Yes --> K[pyvisa.query]
    J -- No  --> L[pyvisa.write]
    K --> M[JSON を stdout に出力]
    L --> M
    M --> N[VBA が stdout を読み取り]
    N --> O[Control シートに\n応答・ステータスを書き込み]

    style A fill:#1F4E79,color:#fff
    style O fill:#1F4E79,color:#fff
```

### 2.2 Flask 方式（通常コマンド）

```mermaid
flowchart TD
    A([Excel ボタン押下]) --> B[GpibControlHttp.bas\nExecuteSelectedCommandHttp]
    B --> C{サーバー\n稼働中?}
    C -- No  --> D[StartGpibServer\nサーバー起動]
    D --> E{起動確認\n/health}
    E -- 失敗 --> F([エラー表示\n終了])
    E -- 成功 --> G
    C -- Yes --> G[Config シートから\nアドレス・タイムアウト取得]
    G --> H{D/E 列に\nProtocol/Host\nあり?}
    H -- あり --> I[AppConfig.BuildVisaAddress\nでアドレス生成]
    H -- なし --> J[B列のアドレスを使用]
    I --> K
    J --> K[POST /execute\nJSON リクエスト送信]
    K --> L[server.py\n/execute エンドポイント]
    L --> M[GpibManager.execute]
    M --> N[接続プールに\nアドレスあり?]
    N -- Yes --> O[既存接続を再利用]
    N -- No  --> P[GenericInstrument\n新規接続 open]
    O --> Q
    P --> Q[instrument.execute\nコマンド送信]
    Q --> R{成功?}
    R -- Yes --> S[JSON レスポンス\n返却]
    R -- No  --> T[リトライ処理\n→ 2.5 参照]
    T --> S
    S --> U[VBA が JSON を解析]
    U --> V[Control シートに\n応答・ステータスを書き込み]

    style A fill:#1F4E79,color:#fff
    style F fill:#CC0000,color:#fff
    style V fill:#1F4E79,color:#fff
```

### 2.3 Flask 方式（MT8821C 専用）

```mermaid
flowchart TD
    A([Excel ボタン押下]) --> B[GpibMT8821C.bas\nまたは MT8821C_Sample.bas]
    B --> C[MT8821C_Call\nアクション名 + パラメータ]
    C --> D[POST /mt8821c/execute\nJSON リクエスト]
    D --> E[mt8821c_blueprint.py\n/mt8821c/execute]
    E --> F{アクション名\nが ACTIONS\nに存在?}
    F -- No  --> G([400 エラー\n未知のアクション])
    F -- Yes --> H[ACTIONS dict から\nSCPI コマンド文字列を生成]
    H --> I[GpibManager.execute\n通常ルートに合流]
    I --> J[pyvisa でコマンド送信]
    J --> K[JSON レスポンス返却]
    K --> L[VBA が解析]
    L --> M{ResultSheet.bas\nインポート済?}
    M -- Yes --> N[Result_AppendRow\nResult シートへ記録]
    M -- No  --> O[MsgBox のみ表示]
    N --> O

    style A fill:#1F4E79,color:#fff
    style G fill:#CC0000,color:#fff
    style N fill:#006400,color:#fff
```

### 2.4 LAN 接続の分岐

```mermaid
flowchart TD
    A[VISAアドレス文字列] --> B{アドレス\nパターン判定}
    B --> C["GPIB0::N::INSTR\n→ GPIB 接続"]
    B --> D["TCPIP0::host::INSTR\n→ VXI-11 接続"]
    B --> E["TCPIP0::host::port::SOCKET\n→ Raw Socket 接続"]
    B --> F["TCPIP0::host::hislip0::INSTR\n→ HiSLIP 接続"]

    C --> G[pyvisa.open_resource\nGPIB ボード経由]
    D --> H[pyvisa.open_resource\nTCP/IP VXI-11 プロトコル]
    E --> I[pyvisa.open_resource\nTCP/IP Raw Socket]
    F --> J[pyvisa.open_resource\nHiSLIP プロトコル]

    I --> K["read_termination = LF\nwrite_termination = LF\n(明示的に設定)"]

    G --> L[BaseInstrument.write / query]
    H --> L
    K --> L
    J --> L

    style E fill:#FFF2CC
    style K fill:#FFF2CC
```

> **注意:** Raw Socket (`::SOCKET`) のみ終端文字 (`\n`) を明示設定する。
> GPIB / VXI-11 / HiSLIP は pyvisa が自動処理する。

### 2.5 リトライ・再接続フロー

```mermaid
flowchart TD
    A[GpibManager.execute 呼び出し] --> B[attempt = 0]
    B --> C[_get_or_create\n接続プール確認]
    C --> D[instrument.execute\nコマンド送信]
    D --> E{成功?}
    E -- Yes --> F([結果を返す])
    E -- No  --> G{attempt <\nmax_retry?}
    G -- No  --> H([エラーを返す])
    G -- Yes --> I[接続をクローズ\n_close_connection]
    I --> J[attempt += 1]
    J --> K[WARNING ログ出力\nRETRY N/max_retry]
    K --> C

    style F fill:#006400,color:#fff
    style H fill:#CC0000,color:#fff
    style K fill:#FF6600,color:#fff
```

> `max_retry` は `config/settings.ini` の `[Server] MaxRetry` で設定 (デフォルト: 1)。

### 2.6 試験結果の記録フロー

```mermaid
flowchart TD
    A([コマンド実行]) --> B{記録方式}

    B --> C[MT8821C_Sample.bas\nから直接記録]
    B --> D[Control シートから\n一括転記]

    C --> E[LogOp 内で\nResult_AppendRow を呼び出し]
    D --> F[Result_AppendFromControl\n実行済み行を全転記]

    E --> G
    F --> G[ResultSheet.bas\nResult_AppendRow]

    G --> H{Result シート\nあり?}
    H -- No  --> I[Result シートを\n自動作成・初期化]
    H -- Yes --> J[最終行の次に追記]
    I --> J
    J --> K[No / 実行日時 / 機器名 /\nアドレス / 接続方式 /\nコマンド / 応答 /\nステータス / 備考\n を書き込む]
    K --> L[ステータスを\n色付き表示\nOK=緑 / ERROR=赤]

    style G fill:#1F4E79,color:#fff
    style L fill:#006400,color:#fff
```

---

## 3. モジュール構成

### Python 側

```mermaid
classDiagram
    class BaseInstrument {
        +address: str
        +connection_type: str
        +open()
        +close()
        +write(command)
        +query(command) str
        +read() str
    }

    class GenericInstrument {
        +execute(command) dict
    }

    class LanInstrument {
        +vxi11(host)$ LanInstrument
        +hislip(host, index)$ LanInstrument
        +socket(host, port)$ LanInstrument
    }

    class AnritsuMT8821C {
        +ACTIONS: dict
        +identify() str
        +reset() str
        +preset() str
        +get_error() str
        +get_dl_power() str
        +set_dl_power(power) str
        +get_band() str
        +set_band(band) str
        +get_channel() str
        +set_channel(channel) str
        +call_connect() str
        +call_disconnect() str
        +get_call_status() str
        +measure_ul_power() str
    }

    class GpibManager {
        -_connections: dict
        -_lock: Lock
        -_max_retry: int
        +execute(address, command, timeout) dict
        +close_connection(address) bool
        +close_all()
        +list_connections() list
        +list_resources() list
        -_get_or_create(address, timeout) GenericInstrument
        -_close_connection(address) bool
    }

    class config {
        +get_server_settings() dict
        +get_lan_settings() dict
        +build_visa_address(protocol, host, port) str
        +setup_logging()
    }

    BaseInstrument <|-- GenericInstrument
    BaseInstrument <|-- LanInstrument
    GenericInstrument <|-- AnritsuMT8821C
    GpibManager --> GenericInstrument : 生成・管理
```

### VBA 側

```
┌───────────────────────────────────────────────────────┐
│ VBA モジュール依存関係                                  │
│                                                       │
│  AppConfig.bas          設定読み込み (必須)             │
│       ↑ 依存                                          │
│  GpibControlHttp.bas    Flask方式 実行  ─┐            │
│  GpibControl.bas        CLI方式 実行    │ Control     │
│  GpibMT8821C.bas        MT8821C専用    ─┘ シート操作   │
│  MT8821C_Sample.bas     動作確認サンプル  ─┐           │
│       ↓ 呼び出し                          │ Result    │
│  ResultSheet.bas        試験結果管理     ←┘ シート操作  │
└───────────────────────────────────────────────────────┘
```

---

## 4. ファイル構成

```
operateGPIBFromVBA/
│
├── python/                          # Python バックエンド
│   ├── server.py                    # Flask サーバー (エンドポイント定義)
│   ├── gpib_manager.py              # 接続プール・リトライ・スレッドセーフ
│   ├── gpib_controller.py           # [CLI] エントリポイント
│   ├── config.py                    # settings.ini 読み込み・ロギング設定
│   ├── requirements.txt             # pyvisa, pyvisa-py, flask
│   ├── instruments/
│   │   ├── base_instrument.py       # 基底クラス (GPIB/LAN 共通)
│   │   ├── generic_instrument.py    # 汎用機器クラス
│   │   ├── lan_instrument.py        # LAN 接続ファクトリクラス
│   │   └── anritsu_mt8821c.py       # MT8821C 専用クラス (ACTIONS dict)
│   └── blueprints/                  # プラグイン Blueprint
│       └── mt8821c_blueprint.py     # MT8821C 専用エンドポイント
│
├── vba/                             # Excel VBA モジュール (SHIFT-JIS)
│   ├── AppConfig.bas                # INI 読み込み・VISA アドレスビルダー
│   ├── GpibControlHttp.bas          # Flask方式 実行・サーバー起動管理
│   ├── GpibControl.bas              # CLI方式 実行
│   ├── GpibMT8821C.bas              # MT8821C 専用操作
│   ├── MT8821C_Sample.bas           # MT8821C 動作確認サンプル集
│   └── ResultSheet.bas              # 試験結果 Result シート管理
│
├── config/
│   └── settings.ini                 # サーバー設定・LAN設定・ログ設定
│
├── docs/
│   └── system_design.md             # 本ファイル (システム設計書)
│
├── tests/                           # pytest テストスイート
│   ├── conftest.py                  # pyvisa モックフィクスチャ
│   ├── test_config.py               # config.py テスト
│   ├── test_instruments.py          # 機器クラステスト
│   ├── test_gpib_manager.py         # GpibManager テスト
│   └── test_mt8821c.py              # MT8821C テスト
│
├── create_excel.py                  # Excel ファイル生成スクリプト
├── debug_client.py                  # CLI デバッグツール
├── start_server.bat                 # サーバー起動バッチ
├── CLAUDE.md                        # エンコーディング規約
└── README.md                        # セットアップ・操作手順
```

---

## 5. Excel シート仕様

### Config シート（機器設定）

| 列 | 項目 | 必須 | 説明 |
|----|------|------|------|
| A | 機器名 (Name) | ✔ | Control シートの A 列と一致させる |
| B | VISA アドレス | △ | フル指定の場合。D/E 列がある場合は自動生成 |
| C | Timeout (ms) | ✔ | 通信タイムアウト。例: 5000 |
| D | Protocol | △ | `GPIB` / `TCPIP` / `SOCKET` / `HISLIP` |
| E | Host / IP | △ | IPアドレスまたはホスト名 |
| F | Port | △ | Raw Socket の場合のみ必要。例: 5025 |

> D/E 列が入力されている場合は B 列より優先して VISA アドレスを自動生成する。

### Control シート（コマンド実行）

| 列 | 項目 | 説明 |
|----|------|------|
| A | 機器名 | Config シートの A 列と一致させる |
| B | SCPI コマンド | `?` で終わると query、それ以外は write |
| C | 応答結果 | 実行後に自動入力される |
| D | ステータス | `OK` (緑) または `ERROR: メッセージ` (赤) |

### Result シート（試験結果記録）

| 列 | 項目 | 説明 |
|----|------|------|
| A | No. | 通し番号 (自動採番) |
| B | 実行日時 | yyyy/mm/dd hh:mm:ss |
| C | 機器名 | 操作対象の機器名 |
| D | VISA アドレス | 接続に使用したアドレス |
| E | 接続方式 | GPIB / LAN VXI-11 / LAN Socket / LAN HiSLIP |
| F | コマンド / アクション | 実行した SCPI コマンドまたはアクション名 |
| G | 応答結果 | 機器からの応答 |
| H | ステータス | OK (緑) / ERROR (赤) |
| I | 備考 | 「Control転記」「比較確認」など |

---

## 6. 設定ファイル仕様（settings.ini）

```ini
[Server]
Host            = 127.0.0.1        # Flask サーバーのバインドアドレス
Port            = 5000             # Flask サーバーのポート番号
PythonExe       = python           # Python 実行ファイルのパス
ServerScript    = C:\...\server.py # server.py の絶対パス
HealthTimeoutSec = 10              # サーバー起動確認のタイムアウト (秒)
MaxRetry        = 1                # コマンド失敗時の再試行回数

[Logging]
Level       = INFO                 # DEBUG / INFO / WARNING / ERROR
LogDir      = logs                 # ログファイルの出力先ディレクトリ
FileName    = gpib.log             # ログファイル名
MaxBytes    = 1000000              # ローテーション上限サイズ (bytes)
BackupCount = 3                    # ローテーション保持世代数
Format      = %(asctime)s [%(levelname)s] %(name)s: %(message)s

[Lan]
DefaultSocketPort  = 5025          # Raw Socket 接続のデフォルトポート
ReadTermination    = \n            # Raw Socket 読み取り終端文字
WriteTermination   = \n            # Raw Socket 書き込み終端文字
```

> `settings.ini` は Python 側 (`configparser`) と VBA 側 (Windows API `GetPrivateProfileString`) の両方から読み込む。

---

## 7. API エンドポイント仕様

### 汎用エンドポイント

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/health` | サーバー稼働確認 |
| POST | `/execute` | SCPI コマンド実行 |
| GET | `/connections` | 現在の接続一覧 |
| DELETE | `/connections/<address>` | 指定接続を閉じる |
| GET | `/resources` | VISA リソース一覧 |
| GET | `/debug` | サーバー内部状態 (デバッグ用) |

**POST /execute リクエスト:**
```json
{ "address": "GPIB0::1::INSTR", "command": "*IDN?", "timeout": 5000 }
```

**POST /execute レスポンス:**
```json
{ "success": true, "response": "ANRITSU,...", "error": "", "address": "...", "command": "..." }
```

### MT8821C 専用エンドポイント（Blueprint）

| メソッド | パス | 説明 |
|---------|------|------|
| POST | `/mt8821c/execute` | 名前付きアクションを実行 |
| GET | `/mt8821c/actions` | 利用可能なアクション一覧 |

**POST /mt8821c/execute リクエスト:**
```json
{ "address": "GPIB0::1::INSTR", "action": "set_dl_power", "params": {"power": -70.0} }
```

**利用可能なアクション一覧 (GET /mt8821c/actions):**

| アクション名 | 対応 SCPI | パラメータ |
|------------|----------|-----------|
| `identify` | `*IDN?` | なし |
| `reset` | `*RST` | なし |
| `preset` | `SYSTem:PRESet` | なし |
| `get_error` | `SYSTem:ERRor?` | なし |
| `get_dl_power` | `BS:OLVL?` | なし |
| `set_dl_power` | `BS:OLVL <power>` | `{"power": float}` |
| `get_band` | `BAND?` | なし |
| `set_band` | `BAND <band>` | `{"band": int}` |
| `get_channel` | `CHANL?` | なし |
| `set_channel` | `CHANL <channel>` | `{"channel": int}` |
| `call_connect` | `CALLSO` | なし |
| `call_disconnect` | `CALLEND` | なし |
| `get_call_status` | `CALLSTAT?` | なし |
| `measure_ul_power` | `MEAS:UL:POW?` | なし |

---

## 8. VISA アドレス仕様

### アドレス形式

| 接続方式 | アドレス形式 | 例 |
|---------|------------|-----|
| GPIB | `GPIB0::<addr>::INSTR` | `GPIB0::1::INSTR` |
| LAN VXI-11 | `TCPIP0::<host>::INSTR` | `TCPIP0::192.168.1.10::INSTR` |
| LAN Raw Socket | `TCPIP0::<host>::<port>::SOCKET` | `TCPIP0::192.168.1.10::5025::SOCKET` |
| LAN HiSLIP | `TCPIP0::<host>::hislip0::INSTR` | `TCPIP0::192.168.1.10::hislip0::INSTR` |

### Config シートの Protocol 列から自動生成するルール

| Protocol 列 (D列) | 生成されるアドレス |
|-----------------|-----------------|
| `GPIB` | `GPIB0::<Host>::INSTR` |
| `TCPIP` / `VXI11` / `LAN` | `TCPIP0::<Host>::INSTR` |
| `SOCKET` / `TCPIP_SOCKET` | `TCPIP0::<Host>::<Port>::SOCKET` (Port 省略時は 5025) |
| `HISLIP` / `TCPIP_HISLIP` | `TCPIP0::<Host>::hislip0::INSTR` |

---

## 9. プラグイン拡張方式

### Blueprint プラグイン（Python 側）

`python/blueprints/` に `*_blueprint.py` を配置するだけで自動ロードされる。

```
python/blueprints/
└── mt8821c_blueprint.py   ← 配置するだけで /mt8821c/* が有効になる
```

```mermaid
flowchart LR
    A[server.py 起動] --> B[_load_blueprints]
    B --> C{blueprints/*.py\nを列挙}
    C --> D["*_blueprint.py\nかつ\n'blueprint' 属性あり?"]
    D -- Yes --> E[app.register_blueprint\n自動登録]
    D -- No  --> F[スキップ]
    E --> G[エンドポイント有効化]
```

**新しい機器クラスの追加手順:**

1. `python/instruments/<機器名>.py` を作成 (`GenericInstrument` を継承)
2. `python/blueprints/<機器名>_blueprint.py` を作成
3. `vba/<機器名>.bas` を作成してインポート

### VBA モジュールプラグイン

VBA モジュールは独立しており、インポート/削除だけで機能の追加/削除ができる。

| モジュール | 役割 | 依存 |
|-----------|------|------|
| `AppConfig.bas` | 設定読み込み | なし (必須) |
| `GpibControlHttp.bas` | Flask方式実行 | AppConfig |
| `GpibControl.bas` | CLI方式実行 | AppConfig |
| `GpibMT8821C.bas` | MT8821C操作 | AppConfig |
| `ResultSheet.bas` | 結果記録 | なし |
| `MT8821C_Sample.bas` | サンプル集 | AppConfig, GpibMT8821C, ResultSheet |
