"""
GPIB Flask サーバー

VBAからHTTP経由でGPIBを制御するためのローカルサーバー。
設定は config/settings.ini から読み込む。

起動方法:
  python server.py
  python server.py --config ../config/settings.ini

エンドポイント:
  GET  /health                  サーバーの稼働確認
  POST /execute                 GPIBコマンドの実行
  GET  /connections             現在オープンしている接続一覧
  DELETE /connections/<address> 指定接続を閉じる
  GET  /resources               VISAで認識されているリソース一覧
"""
import argparse
import atexit
import importlib
import logging
import os
import sys

# pythonフォルダをパスに追加
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, jsonify, request

import config as appconfig
from gpib_manager import GpibManager

app = Flask(__name__)
manager: GpibManager = None
logger = logging.getLogger(__name__)


# ------------------------------------------------------------------
# エンドポイント
# ------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    """サーバーの稼働確認。VBAが起動チェックに使う。"""
    logger.debug("GET /health")
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
        logger.warning("POST /execute: 不正なリクエスト (JSONではない)")
        return jsonify({"success": False, "error": "リクエストボディがJSONではありません", "response": ""}), 400

    address = data.get("address", "").strip()
    command = data.get("command", "").strip()
    timeout = int(data.get("timeout", 5000))

    if not address or not command:
        logger.warning("POST /execute: address または command が未指定")
        return jsonify({"success": False, "error": "address と command は必須です", "response": ""}), 400

    logger.info("RECV | address=%s | command=%s | timeout=%d", address, command, timeout)
    result = manager.execute(address, command, timeout)
    status_code = 200 if result["success"] else 500
    return jsonify(result), status_code


@app.route("/connections", methods=["GET"])
def list_connections():
    """現在オープンしている接続一覧"""
    connections = manager.list_connections()
    logger.debug("GET /connections: %d 件", len(connections))
    return jsonify({"connections": connections})


@app.route("/connections/<path:address>", methods=["DELETE"])
def close_connection(address: str):
    """指定アドレスの接続を閉じる"""
    logger.info("DELETE /connections/%s", address)
    closed = manager.close_connection(address)
    if closed:
        return jsonify({"success": True, "address": address})
    return jsonify({"success": False, "error": f"接続が見つかりません: {address}"}), 404


@app.route("/resources", methods=["GET"])
def list_resources():
    """VISAで認識されているリソース一覧"""
    resources = manager.list_resources()
    logger.debug("GET /resources: %s", resources)
    return jsonify({"resources": resources})


@app.route("/debug", methods=["GET"])
def debug_info():
    """デバッグ用: サーバー内部状態をまとめて返す"""
    import sys
    import logging
    connections = manager.list_connections()
    resources = manager.list_resources()
    root_level = logging.getLogger().level
    return jsonify({
        "server": {
            "python_version": sys.version,
            "log_level": logging.getLevelName(root_level),
        },
        "connections": {
            "count": len(connections),
            "addresses": connections,
        },
        "resources": {
            "count": len(resources),
            "list": resources,
        },
        "blueprints": [bp for bp in app.blueprints],
    })


# ------------------------------------------------------------------
# Blueprint 自動ロード
# ------------------------------------------------------------------

def _load_blueprints(app: Flask) -> None:
    """
    blueprints/ ディレクトリにある *_blueprint.py を自動検出して登録する。

    【追加】 blueprints/ に *_blueprint.py を置くだけで有効になる
    【削除】 blueprints/ からファイルを削除するだけで無効になる
    """
    bp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blueprints")
    if not os.path.isdir(bp_dir):
        return
    for fname in sorted(os.listdir(bp_dir)):
        if not fname.endswith("_blueprint.py") or fname.startswith("_"):
            continue
        module_name = f"blueprints.{fname[:-3]}"
        try:
            module = importlib.import_module(module_name)
            if hasattr(module, "blueprint"):
                app.register_blueprint(module.blueprint)
                logger.info("Blueprint 登録: %s (prefix=%s)", fname, module.blueprint.url_prefix)
        except Exception as e:
            logger.error("Blueprint ロード失敗: %s : %s", fname, e)


# ------------------------------------------------------------------
# サーバー起動
# ------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(description="GPIB Flask Server")
    parser.add_argument("--config", default=None, help="settings.ini のパス (省略時はデフォルトパスを使用)")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    # 設定読み込みとロギング初期化
    appconfig.setup_logging(args.config)
    server_settings = appconfig.get_server_settings(args.config)

    # GpibManager 初期化 (シャットダウン時にすべての接続を閉じる)
    manager = GpibManager(max_retry=server_settings["max_retry"])
    atexit.register(manager.close_all)

    # Blueprint から参照できるよう app に紐付ける
    app.gpib_manager = manager
    _load_blueprints(app)

    logger.info(
        "GPIB Server 起動 | http://%s:%d | max_retry=%d",
        server_settings["host"], server_settings["port"], server_settings["max_retry"],
    )

    app.run(
        host=server_settings["host"],
        port=server_settings["port"],
        debug=False,
        use_reloader=False,
    )
