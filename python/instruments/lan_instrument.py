"""
LAN (TCP/IP) 接続の計測器制御クラス

ホスト名/IPアドレスと接続方式から VISA アドレスを自動生成するファクトリメソッドを提供する。
GPIB との違いは VISA アドレスの形式のみ。SCPI コマンドの送受信は BaseInstrument と共通。

接続方式と VISA アドレス:
  VXI-11:  TCPIP0::<host>::INSTR            多くの計測器の標準 LAN 制御
  HiSLIP:  TCPIP0::<host>::hislip0::INSTR   高速通信 (Keysight 等)
  Socket:  TCPIP0::<host>::<port>::SOCKET   Raw TCP。port は機器依存 (通常 5025)
"""
from .generic_instrument import GenericInstrument

# Raw Socket 接続のデフォルトポート (SCPI over TCP の業界標準)
DEFAULT_SOCKET_PORT = 5025


class LanInstrument(GenericInstrument):
    """LAN (TCP/IP) 接続の計測器制御クラス"""

    # ---- ファクトリメソッド ------------------------------------------------

    @classmethod
    def vxi11(cls, host: str, timeout: int = 5000) -> "LanInstrument":
        """VXI-11 接続インスタンスを生成する (最も汎用的な LAN 制御方式)

        Args:
            host: ホスト名または IP アドレス (例: "192.168.1.100")
            timeout: タイムアウト(ms)
        """
        return cls(address=f"TCPIP0::{host}::INSTR", timeout=timeout)

    @classmethod
    def hislip(cls, host: str, index: int = 0, timeout: int = 5000) -> "LanInstrument":
        """HiSLIP 接続インスタンスを生成する (高速 LAN 制御)

        Args:
            host: ホスト名または IP アドレス
            index: HiSLIP インターフェース番号 (通常 0)
            timeout: タイムアウト(ms)
        """
        return cls(address=f"TCPIP0::{host}::hislip{index}::INSTR", timeout=timeout)

    @classmethod
    def socket(cls, host: str, port: int = DEFAULT_SOCKET_PORT, timeout: int = 5000) -> "LanInstrument":
        """Raw Socket 接続インスタンスを生成する

        Args:
            host: ホスト名または IP アドレス
            port: TCP ポート番号 (デフォルト: 5025)
            timeout: タイムアウト(ms)
        """
        return cls(address=f"TCPIP0::{host}::{port}::SOCKET", timeout=timeout)
