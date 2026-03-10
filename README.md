# operateGPIBFromVBA

Excel VBA から Python (pyvisa) を経由して GPIB 機器を制御するシステム。

## 構成

```
operateGPIBFromVBA/
├── python/
│   ├── gpib_controller.py        # [CLI方式] エントリポイント
│   ├── server.py                 # [Flask方式] HTTPサーバー (安定運用向け)
│   ├── gpib_manager.py           # 接続プール・リトライ・ログ管理
│   ├── instruments/
│   │   ├── base_instrument.py    # 基底クラス
│   │   └── generic_instrument.py # 汎用機器クラス
│   └── requirements.txt
├── vba/
│   ├── GpibControl.bas           # [CLI方式] VBAモジュール
│   └── GpibControlHttp.bas       # [Flask方式] VBAモジュール (安定運用向け)
├── start_server.bat              # Flaskサーバー起動バッチ
└── README.md
```

## 2つの方式

| 項目 | CLI方式 (`GpibControl.bas`) | Flask方式 (`GpibControlHttp.bas`) |
|------|----------------------------|------------------------------------|
| Pythonプロセス | コマンド毎に起動 | サーバーとして常駐 |
| GPIB接続 | 毎回 open/close | 接続プールで再利用 |
| リトライ | なし | 失敗時に自動再接続+リトライ |
| ログ | なし | ファイル+コンソールに記録 |
| 用途 | 試験・デバッグ | 本番運用 |

---

## セットアップ

### 1. Python環境

```bash
pip install -r python/requirements.txt
```

> **注意:** NI-VISA がインストール済みの場合は `pyvisa` だけで動作する。
> NI-VISA がない場合は `pyvisa-py` を使うが、別途GPIBバックエンドが必要。

### 2. VBAモジュールのインポート

1. Excel で `Alt + F11` → VBAエディタを開く
2. `ファイル` → `ファイルのインポート` で以下を選択:
   - CLI方式: `vba/GpibControl.bas`
   - Flask方式: `vba/GpibControlHttp.bas` (推奨)
3. 各モジュール内の定数を環境に合わせて変更する

**GpibControl.bas (CLI方式):**
```vba
Private Const PYTHON_EXE As String = "python"
Private Const SCRIPT_PATH As String = "C:\work\operateGPIBFromVBA\operateGPIBFromVBA\python\gpib_controller.py"
```

**GpibControlHttp.bas (Flask方式):**
```vba
Private Const SERVER_BASE_URL As String = "http://127.0.0.1:5000"
Private Const SERVER_SCRIPT As String = "C:\work\operateGPIBFromVBA\operateGPIBFromVBA\python\server.py"
```

### 3. Excel シートの準備

#### Config シート (機器設定)

| A: Name      | B: Address             | C: Timeout |
|--------------|------------------------|------------|
| DMM          | GPIB0::22::INSTR       | 5000       |
| PowerSupply  | GPIB0::5::INSTR        | 3000       |

#### Control シート (操作画面)

| A: 機器名    | B: SCPIコマンド | C: 応答結果 | D: ステータス |
|-------------|----------------|------------|--------------|
| DMM         | *IDN?          |            |              |
| PowerSupply | VOLT 5.0       |            |              |

### 4. ボタンの作成

Control シートにボタン（フォームコントロール）を挿入し、マクロを割り当てる。

**CLI方式:**
| ボタン名 | 割り当てるマクロ |
|---------|----------------|
| 選択行を実行 | `GpibControl.ExecuteSelectedCommand` |
| すべて実行 | `GpibControl.ExecuteAllCommands` |

**Flask方式 (推奨):**
| ボタン名 | 割り当てるマクロ |
|---------|----------------|
| サーバー起動 | `GpibControlHttp.StartGpibServer` |
| 選択行を実行 | `GpibControlHttp.ExecuteSelectedCommandHttp` |
| すべて実行 | `GpibControlHttp.ExecuteAllCommandsHttp` |

---

## 処理フロー

### CLI方式
```
[Excelボタン押下]
      ↓
[VBA: Configシートからアドレス取得]
      ↓
[VBA: WScript.Shell.Exec() でPythonを起動]
      ↓
[Python: pyvisa でGPIBコマンドを送信]
      ↓
[Python: JSON を stdout に出力]
      ↓
[VBA: stdout を読み取り、Controlシートに書き込む]
```

### Flask方式 (安定運用)
```
[start_server.bat でサーバーを起動] ← 初回のみ
      ↓
[Excelボタン押下]
      ↓
[VBA: Configシートからアドレス取得]
      ↓
[VBA: POST http://localhost:5000/execute]
      ↓
[Flask: GpibManager が接続プールから接続を取得]
      ↓
[pyvisa: GPIBコマンドを送信 (失敗時は再接続してリトライ)]
      ↓
[Flask: JSON レスポンスを返す]
      ↓
[VBA: レスポンスを解析し、Controlシートに書き込む]
```

---

## Flask サーバー APIエンドポイント

| Method | URL | 説明 |
|--------|-----|------|
| GET | `/health` | サーバー稼働確認 |
| POST | `/execute` | GPIBコマンド実行 |
| GET | `/connections` | 現在の接続一覧 |
| DELETE | `/connections/<address>` | 指定接続を閉じる |
| GET | `/resources` | VISAリソース一覧 |

**POST /execute リクエスト例:**
```json
{"address": "GPIB0::1::INSTR", "command": "*IDN?", "timeout": 5000}
```

**レスポンス例:**
```json
{"success": true, "response": "Keysight,34461A,...", "error": "", "address": "GPIB0::1::INSTR", "command": "*IDN?"}
```

---

## コマンドラインからのテスト

**CLI方式:**
```bash
python python/gpib_controller.py --address "GPIB0::1::INSTR" --command "*IDN?"
python python/gpib_controller.py --address "GPIB0::1::INSTR" --command "VOLT 5.0" --timeout 3000
```

**Flask方式:**
```bash
# サーバー起動
python python/server.py --port 5000

# 別ターミナルでテスト
curl -X POST http://localhost:5000/execute -H "Content-Type: application/json" -d "{\"address\":\"GPIB0::1::INSTR\",\"command\":\"*IDN?\"}"
```

---

## 拡張方法

### 機器固有クラスの追加

`instruments/base_instrument.py` を継承して機器クラスを作成する:

```python
# instruments/power_supply.py
from .base_instrument import BaseInstrument

class PowerSupply(BaseInstrument):
    def set_voltage(self, voltage: float):
        self.write(f"VOLT {voltage:.3f}")

    def get_voltage(self) -> float:
        return float(self.query("VOLT?"))
```

### Flask サーバーへのエンドポイント追加

`server.py` に新しいルートを追加するだけでよい:

```python
@app.route("/voltage", methods=["POST"])
def set_voltage():
    data = request.get_json()
    result = manager.execute(data["address"], f"VOLT {data['voltage']:.3f}")
    return jsonify(result)
```
