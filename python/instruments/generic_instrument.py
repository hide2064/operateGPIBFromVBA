"""
汎用GPIB機器クラス
機器種別を問わず、任意のSCPIコマンドを送受信できる
"""
import logging

from .base_instrument import BaseInstrument

logger = logging.getLogger(__name__)


class GenericInstrument(BaseInstrument):
    """汎用GPIB機器クラス"""

    def identify(self) -> str:
        """機器の識別情報を取得する (*IDN?)"""
        return self.query("*IDN?")

    def reset(self):
        """機器をリセットする (*RST)"""
        self.write("*RST")

    def clear(self):
        """ステータスレジスタをクリアする (*CLS)"""
        self.write("*CLS")

    def execute(self, command: str) -> dict:
        """
        コマンドを実行する
        '?' で終わるコマンドはqueryとして応答を取得し、それ以外はwriteとして送信する

        Returns:
            {"success": bool, "response": str, "error": str}
        """
        result = {"success": False, "response": "", "error": ""}
        is_query = command.strip().endswith("?")
        exec_type = "QUERY" if is_query else "WRITE"
        logger.debug("EXEC  %s addr=%s cmd=%s", exec_type, self._address, command)
        try:
            if is_query:
                result["response"] = self.query(command)
            else:
                self.write(command)
                result["response"] = "OK"
            result["success"] = True
        except Exception as e:
            result["error"] = str(e)
            logger.debug("EXEC  %s FAIL addr=%s cmd=%s error=%s", exec_type, self._address, command, e)
        return result
