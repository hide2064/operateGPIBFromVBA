"""
GPIB Controller - VBAから呼び出されるエントリポイント

使用方法:
  python gpib_controller.py --address <VISAアドレス> --command <SCPIコマンド> [--timeout <ms>]

例:
  python gpib_controller.py --address "GPIB0::1::INSTR" --command "*IDN?"
  python gpib_controller.py --address "GPIB0::2::INSTR" --command "VOLT 5.0" --timeout 3000

出力:
  JSON形式でstdoutに出力する
  {"success": true/false, "response": "...", "error": "..."}
"""
import sys
import json
import argparse

# pythonフォルダをパスに追加
sys.path.insert(0, __file__.rsplit("\\", 1)[0] if "\\" in __file__ else __file__.rsplit("/", 1)[0])

from instruments.generic_instrument import GenericInstrument


def parse_args():
    parser = argparse.ArgumentParser(description="GPIB Controller")
    parser.add_argument("--address", required=True, help='VISAリソースアドレス (例: "GPIB0::1::INSTR")')
    parser.add_argument("--command", required=True, help="送信するSCPIコマンド")
    parser.add_argument("--timeout", type=int, default=5000, help="タイムアウト(ms)、デフォルト: 5000")
    return parser.parse_args()


def main():
    args = parse_args()
    result = {"success": False, "response": "", "error": ""}

    try:
        with GenericInstrument(address=args.address, timeout=args.timeout) as instrument:
            result = instrument.execute(args.command)
    except Exception as e:
        result["error"] = str(e)

    # 結果をJSONでstdoutに出力 (VBAが読み取る)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["success"] else 1


if __name__ == "__main__":
    sys.exit(main())
