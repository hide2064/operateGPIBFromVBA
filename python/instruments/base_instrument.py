"""
GPIB機器の基底クラス
すべての機器クラスはこのクラスを継承する
"""
import pyvisa


class BaseInstrument:
    """GPIB機器の基底クラス"""

    def __init__(self, address: str, timeout: int = 5000):
        """
        Args:
            address: VISAリソースアドレス (例: "GPIB0::1::INSTR")
            timeout: タイムアウト(ms)
        """
        self._address = address
        self._timeout = timeout
        self._rm = pyvisa.ResourceManager()
        self._instrument = None

    def open(self):
        """機器との接続を開く"""
        self._instrument = self._rm.open_resource(self._address)
        self._instrument.timeout = self._timeout

    def close(self):
        """機器との接続を閉じる"""
        if self._instrument:
            self._instrument.close()
            self._instrument = None

    def write(self, command: str):
        """コマンドを送信する (応答なし)"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        self._instrument.write(command)

    def query(self, command: str) -> str:
        """コマンドを送信し、応答を受信する"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        return self._instrument.query(command).strip()

    def read(self) -> str:
        """応答を受信する"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        return self._instrument.read().strip()

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @property
    def address(self) -> str:
        return self._address
