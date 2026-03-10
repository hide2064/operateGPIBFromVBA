"""
Anritsu MT8821C Radio Communication Analyzer 制御クラス

【機器概要】
  LTE / 5G NR 等の無線通信規格に対応したラジオコミュニケーションアナライザ。
  基地局シミュレーション・プロトコル試験・RF 測定に使用する。

【参照マニュアル】
  MT8821C Radio Communication Analyzer Operation Manual
  ※ 以下のコマンドは一般的な Anritsu SCPI 規則に基づく。
     実機環境で動作確認し、必要に応じてコマンド構文を修正すること。

【アーキテクチャ上の役割】
  - GenericInstrument を継承し、MT8821C 固有の名前付きメソッドを提供する
  - ACTIONS クラス変数でアクション定義を一元管理し、Flask Blueprint と共有する
  - CLI / スタンドアロン用途でも単独で使用できる
"""
from .generic_instrument import GenericInstrument


class AnritsuMT8821C(GenericInstrument):
    """Anritsu MT8821C Radio Communication Analyzer"""

    # ---- アクション定義 (Blueprint と共有) --------------------------------
    # 書式: "アクション名": (コマンド生成関数(params) -> str, 説明文)
    # params は dict。コマンド生成関数は必要なキーを取り出して SCPI 文字列を返す。
    ACTIONS: dict[str, tuple] = {
        # --- システム ---
        "identify": (
            lambda p: "*IDN?",
            "機器識別情報の取得 (*IDN?)",
        ),
        "reset": (
            lambda p: "*RST",
            "リセット (*RST)",
        ),
        "preset": (
            lambda p: "SYSTem:PRESet",
            "システムプリセット (SYSTem:PRESet)",
        ),
        "get_error": (
            lambda p: "SYSTem:ERRor?",
            "エラー情報の取得 (SYSTem:ERRor?)",
        ),
        # --- セル設定 (LTE/5GNR 共通) ---
        "get_dl_power": (
            lambda p: "BS:OLVL?",
            "ダウンリンク出力レベルの取得 (BS:OLVL?)",
        ),
        "set_dl_power": (
            lambda p: f"BS:OLVL {float(p['power']):.1f}",
            "ダウンリンク出力レベルの設定 / params: {power: dBm値 (例: -70.0)}",
        ),
        "get_band": (
            lambda p: "BAND?",
            "バンドの取得 (BAND?)",
        ),
        "set_band": (
            lambda p: f"BAND {int(p['band'])}",
            "バンドの設定 / params: {band: バンド番号 (例: 1)}",
        ),
        "get_channel": (
            lambda p: "CHANL?",
            "チャネル番号の取得 (CHANL?)",
        ),
        "set_channel": (
            lambda p: f"CHANL {int(p['channel'])}",
            "チャネル番号の設定 / params: {channel: チャネル番号 (例: 300)}",
        ),
        # --- コール処理 ---
        "call_connect": (
            lambda p: "CALLSO",
            "コール接続開始 (CALLSO: Call Start Over)",
        ),
        "call_disconnect": (
            lambda p: "CALLEND",
            "コール切断 (CALLEND)",
        ),
        "get_call_status": (
            lambda p: "CALLSTAT?",
            "コール状態の取得 (CALLSTAT?)",
        ),
        # --- 測定 ---
        "measure_ul_power": (
            lambda p: "MEAS:UL:POW?",
            "アップリンク電力の測定 (MEAS:UL:POW?)",
        ),
    }

    # ---- 名前付きメソッド (CLI / スタンドアロン用途) ----------------------

    def identify(self) -> dict:
        return self.execute("*IDN?")

    def reset(self) -> dict:
        return self.execute("*RST")

    def preset(self) -> dict:
        """システムプリセット"""
        return self.execute("SYSTem:PRESet")

    def get_error(self) -> dict:
        """エラー情報を取得"""
        return self.execute("SYSTem:ERRor?")

    def get_dl_power(self) -> dict:
        """ダウンリンク出力レベルを取得 (dBm)"""
        return self.execute("BS:OLVL?")

    def set_dl_power(self, power_dbm: float) -> dict:
        """ダウンリンク出力レベルを設定 (dBm)"""
        return self.execute(f"BS:OLVL {power_dbm:.1f}")

    def get_band(self) -> dict:
        """バンドを取得"""
        return self.execute("BAND?")

    def set_band(self, band: int) -> dict:
        """バンドを設定"""
        return self.execute(f"BAND {band}")

    def get_channel(self) -> dict:
        """チャネル番号を取得"""
        return self.execute("CHANL?")

    def set_channel(self, channel: int) -> dict:
        """チャネル番号を設定"""
        return self.execute(f"CHANL {channel}")

    def call_connect(self) -> dict:
        """コール接続を開始する"""
        return self.execute("CALLSO")

    def call_disconnect(self) -> dict:
        """コールを切断する"""
        return self.execute("CALLEND")

    def get_call_status(self) -> dict:
        """コール状態を取得する"""
        return self.execute("CALLSTAT?")

    def measure_ul_power(self) -> dict:
        """アップリンク電力を測定する"""
        return self.execute("MEAS:UL:POW?")
