"""
GPIB接続管理クラス

設計思想 (安定性優先):
  - 接続プール: 一度開いた接続を再利用し、毎回の open/close を避ける
  - 自動再接続: コマンド失敗時に接続を張り直して1回リトライする
  - スレッドセーフ: Flaskの並行リクエストに対応するためロックを使用
  - ロギング: 送受信・エラーをファイルとコンソールに記録
"""
import logging
import threading
import time
from typing import Optional

import pyvisa

from instruments.generic_instrument import GenericInstrument

# ===== ロガー設定 =====
logger = logging.getLogger(__name__)


class GpibManager:
    """GPIB接続プールと実行管理を担うクラス"""

    def __init__(self, max_retry: int = 1):
        """
        Args:
            max_retry: コマンド失敗時の再試行回数 (デフォルト1回)
        """
        self._connections: dict[str, GenericInstrument] = {}
        self._lock = threading.Lock()
        self._max_retry = max_retry
        self._rm = pyvisa.ResourceManager()

    # ------------------------------------------------------------------
    # パブリックメソッド
    # ------------------------------------------------------------------

    def execute(self, address: str, command: str, timeout: int = 5000) -> dict:
        """
        指定アドレスの機器にコマンドを送信する

        失敗時は接続を張り直して max_retry 回リトライする。

        Returns:
            {"success": bool, "response": str, "error": str, "address": str, "command": str}
        """
        result = {
            "success": False,
            "response": "",
            "error": "",
            "address": address,
            "command": command,
        }

        t_start = time.perf_counter()
        last_error = ""
        for attempt in range(self._max_retry + 1):
            try:
                instrument = self._get_or_create(address, timeout)
                cmd_result = instrument.execute(command)
                result.update(cmd_result)
                if result["success"]:
                    elapsed = int((time.perf_counter() - t_start) * 1000)
                    logger.info(
                        "OK    addr=%s cmd=%s resp=%s elapsed=%dms",
                        address, command, result["response"], elapsed,
                    )
                    return result
                last_error = result["error"]
            except Exception as e:
                last_error = str(e)

            # 失敗した場合は接続を閉じて再接続を促す
            if attempt < self._max_retry:
                logger.warning(
                    "RETRY (%d/%d) addr=%s cmd=%s error=%s",
                    attempt + 1, self._max_retry, address, command, last_error,
                )
                self._close_connection(address)

        elapsed = int((time.perf_counter() - t_start) * 1000)
        result["success"] = False
        result["error"] = last_error
        logger.error(
            "FAIL  addr=%s cmd=%s error=%s elapsed=%dms",
            address, command, last_error, elapsed,
        )
        return result

    def close_connection(self, address: str) -> bool:
        """指定アドレスの接続を明示的に閉じる"""
        return self._close_connection(address)

    def close_all(self):
        """すべての接続を閉じる (サーバーシャットダウン時に呼ぶ)"""
        with self._lock:
            addresses = list(self._connections.keys())
        for address in addresses:
            self._close_connection(address)
        logger.info("すべての接続を閉じました")

    def list_connections(self) -> list[dict]:
        """現在オープンしている接続一覧を返す"""
        with self._lock:
            return [{"address": addr} for addr in self._connections]

    def list_resources(self) -> list[str]:
        """VISAで認識されているリソース一覧を返す"""
        try:
            return list(self._rm.list_resources())
        except Exception as e:
            logger.error("list_resources エラー: %s", e)
            return []

    # ------------------------------------------------------------------
    # プライベートメソッド
    # ------------------------------------------------------------------

    def _get_or_create(self, address: str, timeout: int) -> GenericInstrument:
        """接続プールから取得、なければ新規作成して接続する"""
        with self._lock:
            if address in self._connections:
                logger.debug("POOL  HIT  addr=%s (接続再利用)", address)
                return self._connections[address]

            logger.debug("POOL  MISS addr=%s (新規接続を作成)", address)
            instrument = GenericInstrument(address=address, timeout=timeout)
            instrument.open()
            self._connections[address] = instrument
            logger.info(
                "POOL  ADD  addr=%s type=%s (接続プールに追加 / 現在 %d 件)",
                address, instrument.connection_type, len(self._connections),
            )
            return instrument

    def _close_connection(self, address: str) -> bool:
        """接続を閉じてプールから削除する"""
        with self._lock:
            instrument = self._connections.pop(address, None)
        if instrument:
            try:
                instrument.close()
                logger.info(
                    "POOL  DEL  addr=%s (接続プールから削除 / 残 %d 件)",
                    address, len(self._connections),
                )
                return True
            except Exception as e:
                logger.warning("CLOSE FAIL addr=%s error=%s", address, e)
        return False
