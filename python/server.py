"""
GPIB Flask サーバー

VBAからHTTP経由でGPIBを制御するためのローカルサーバー。
GpibManager を介して接続プール・リトライ・ロギングを提供する。

起動方法:
  python server.py
  python server.py --port 5000 --log-level INFO

エンドポイント:
  GET  /health                  サーバーの稼働確認
  POST /execute                 GPIBコマンドの実行
  GET  /connections             現在オープンしている接続一覧
  DELETE /connections/<address> 指定接続を閉じる
  GET  /resources               VISAで認識されているリソース一覧
"""
import argparse
import atexit
import logging
import logging.handlers
import os
import sys

# pythonフォルダをパスに追加 (server.py と同じ階層の instruments を import するため)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, jsonify, request

from gpib_manager import GpibManager

# ===== アプリケーション =====
app = Flask(__name__)
manager: GpibManager = None  # setup_manager() で初期化


# ------------------------------------------------------------------
# エンドポイント
# ------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    """サーバーの稼働確認。VBAが起動チェックに使う。"""
    return jsonify({"status": "ok"})


@app.route("/execute", methods=["POST"])
def execute():
    """
    GPIBコマンドを実行する

    Request JSON:
        {"address": "GPIB0::1::INSTR", "command": "*IDN?", "timeout": 5000}

    Response JSON:
        {"success": true, "response": "...", "error": "", "address": "...", "command": "..."}
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "リクエストボディがJSONではありません", "response": ""}), 400

    address = data.get("address", "").strip()
    command = data.get("command", "").strip()
    timeout = int(data.get("timeout", 5000))

    if not address or not command:
        return jsonify({"success": False, "error": "address と command は必須です", "response": ""}), 400

    result = manager.execute(address, command, timeout)
    status_code = 200 if result["success"] else 500
    return jsonify(result), status_code


@app.route("/connections", methods=["GET"])
def list_connections():
    """現在オープンしている接続一覧"""
    return jsonify({"connections": manager.list_connections()})


@app.route("/connections/<path:address>", methods=["DELETE"])
def close_connection(address: str):
    """指定アドレスの接続を閉じる"""
    closed = manager.close_connection(address)
    if closed:
        return jsonify({"success": True, "address": address})
    return jsonify({"success": False, "error": f"接続が見つかりません: {address}"}), 404


@app.route("/resources", methods=["GET"])
def list_resources():
    """VISAで認識されているリソース一覧"""
    return jsonify({"resources": manager.list_resources()})


# ------------------------------------------------------------------
# サーバー起動
# ------------------------------------------------------------------

def setup_logging(log_level: str, log_file: str):
    """ロギング設定: コンソール + ローテーションファイル"""
    level = getattr(logging, log_level.upper(), logging.INFO)
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

    handlers = [logging.StreamHandler()]
    if log_file:
        os.makedirs(os.path.dirname(os.path.abspath(log_file)), exist_ok=True)
        handlers.append(
            logging.handlers.RotatingFileHandler(
                log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
            )
        )

    logging.basicConfig(level=level, format=fmt, handlers=handlers)


def setup_manager(max_retry: int) -> GpibManager:
    """GpibManagerを初期化し、シャットダウン時のクリーンアップを登録する"""
    m = GpibManager(max_retry=max_retry)
    atexit.register(m.close_all)
    return m


def parse_args():
    parser = argparse.ArgumentParser(description="GPIB Flask Server")
    parser.add_argument("--port", type=int, default=5000, help="ポート番号 (デフォルト: 5000)")
    parser.add_argument("--host", default="127.0.0.1", help="バインドアドレス (デフォルト: 127.0.0.1)")
    parser.add_argument("--max-retry", type=int, default=1, help="コマンド失敗時のリトライ回数 (デフォルト: 1)")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    parser.add_argument("--log-file", default="logs/gpib_server.log", help="ログファイルパス")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    setup_logging(args.log_level, args.log_file)
    manager = setup_manager(args.max_retry)

    logging.getLogger(__name__).info(
        "GPIB Serverを起動します: http://%s:%d", args.host, args.port
    )
    # debug=False, use_reloader=False で安定稼働させる
    app.run(host=args.host, port=args.port, debug=False, use_reloader=False)
