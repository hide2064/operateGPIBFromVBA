"""
計測器の基底クラス
GPIB / LAN (VXI-11 / HiSLIP / Raw Socket) のいずれの接続方式でも動作する。
すべての機器クラスはこのクラスを継承する。
"""
import logging

import pyvisa

logger = logging.getLogger(__name__)


class BaseInstrument:
    """計測器の基底クラス"""

    def __init__(self, address: str, timeout: int = 5000):
        """
        Args:
            address: VISAリソースアドレス
                GPIB 例: "GPIB0::1::INSTR"
                LAN VXI-11 例: "TCPIP0::192.168.1.1::INSTR"
                LAN Socket 例: "TCPIP0::192.168.1.1::5025::SOCKET"
                LAN HiSLIP 例: "TCPIP0::192.168.1.1::hislip0::INSTR"
            timeout: タイムアウト(ms)
        """
        self._address = address
        self._timeout = timeout
        self._rm = pyvisa.ResourceManager()
        self._instrument = None

    def open(self):
        """機器との接続を開く"""
        conn_type = self.connection_type
        logger.debug("OPEN  [%s] addr=%s timeout=%dms", conn_type, self._address, self._timeout)
        self._instrument = self._rm.open_resource(self._address)
        self._instrument.timeout = self._timeout
        # Raw Socket 接続は終端文字が自動設定されないため明示的に設定する
        # (GPIB / VXI-11 / HiSLIP は pyvisa が自動処理する)
        if "::SOCKET" in self._address.upper():
            self._instrument.read_termination  = "\n"
            self._instrument.write_termination = "\n"
            logger.debug("TERM  [%s] read_termination=LF write_termination=LF (Socket用に明示設定)", conn_type)
        logger.info("OPEN  [%s] addr=%s — 接続完了", conn_type, self._address)

    def close(self):
        """機器との接続を閉じる"""
        if self._instrument:
            logger.debug("CLOSE [%s] addr=%s", self.connection_type, self._address)
            self._instrument.close()
            self._instrument = None

    def write(self, command: str):
        """コマンドを送信する (応答なし)"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        logger.debug("WRITE [%s] addr=%s cmd=%s", self.connection_type, self._address, command)
        self._instrument.write(command)

    def query(self, command: str) -> str:
        """コマンドを送信し、応答を受信する"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        logger.debug("QUERY [%s] addr=%s cmd=%s", self.connection_type, self._address, command)
        resp = self._instrument.query(command).strip()
        logger.debug("RESP  [%s] addr=%s resp=%s", self.connection_type, self._address, resp)
        return resp

    def read(self) -> str:
        """応答を受信する"""
        if not self._instrument:
            raise RuntimeError("機器が接続されていません。open()を先に呼び出してください。")
        resp = self._instrument.read().strip()
        logger.debug("READ  [%s] addr=%s resp=%s", self.connection_type, self._address, resp)
        return resp

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @property
    def address(self) -> str:
        return self._address

    @property
    def connection_type(self) -> str:
        """接続方式を返す (GPIB / TCPIP_VXI11 / TCPIP_SOCKET / TCPIP_HISLIP / UNKNOWN)"""
        addr_upper = self._address.upper()
        if addr_upper.startswith("GPIB"):
            return "GPIB"
        if "::SOCKET" in addr_upper:
            return "TCPIP_SOCKET"
        if "HISLIP" in addr_upper:
            return "TCPIP_HISLIP"
        if addr_upper.startswith("TCPIP"):
            return "TCPIP_VXI11"
        return "UNKNOWN"
