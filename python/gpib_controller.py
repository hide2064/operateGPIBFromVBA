"""
GPIB Controller - VBAから呼び出されるエントリポイント (CLI方式)

使用方法:
  python gpib_controller.py --address <VISAアドレス> --command <SCPIコマンド> [オプション]

オプション:
  --timeout <ms>    タイムアウト(ミリ秒)。デフォルトは settings.ini の値を使用。
  --config  <path>  settings.ini のパス。省略時はデフォルトパスを使用。

例:
  python gpib_controller.py --address "GPIB0::1::INSTR" --command "*IDN?"
  python gpib_controller.py --address "GPIB0::2::INSTR" --command "VOLT 5.0" --timeout 3000

出力:
  JSON形式でstdoutに出力する
  {"success": true/false, "response": "...", "error": "..."}

ログ:
  settings.ini の [Logging] 設定に従い、ファイルとコンソールに出力する。
"""
import argparse
import json
import logging
import os
import sys

# pythonフォルダをパスに追加
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config as appconfig
from instruments.generic_instrument import GenericInstrument

logger = logging.getLogger(__name__)


def parse_args():
    parser = argparse.ArgumentParser(description="GPIB Controller (CLI)")
    parser.add_argument("--address", required=True, help='VISAリソースアドレス (例: "GPIB0::1::INSTR")')
    parser.add_argument("--command", required=True, help="送信するSCPIコマンド")
    parser.add_argument("--timeout", type=int, default=None, help="タイムアウト(ms)。省略時は settings.ini の値を使用")
    parser.add_argument("--config",  default=None, help="settings.ini のパス")
    return parser.parse_args()


def main():
    args = parse_args()

    # ロギングと設定の初期化
    appconfig.setup_logging(args.config)
    server_settings = appconfig.get_server_settings(args.config)
    timeout = args.timeout if args.timeout is not None else 5000

    result = {"success": False, "response": "", "error": ""}

    logger.info("CLI実行開始 | address=%s | command=%s | timeout=%d", args.address, args.command, timeout)

    try:
        with GenericInstrument(address=args.address, timeout=timeout) as instrument:
            result = instrument.execute(args.command)

        if result["success"]:
            logger.info("CLI実行成功 | response=%s", result["response"])
        else:
            logger.error("CLI実行失敗 | error=%s", result["error"])

    except Exception as e:
        result["error"] = str(e)
        logger.exception("CLI実行中に例外が発生 | address=%s | command=%s", args.address, args.command)

    # 結果をJSONでstdoutに出力 (VBAが読み取る)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["success"] else 1


if __name__ == "__main__":
    sys.exit(main())
