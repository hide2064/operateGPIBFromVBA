"""
Anritsu MT8821C 専用 Flask Blueprint

【追加方法】 このファイルを python/blueprints/ に置くだけで自動ロードされる
【削除方法】 このファイルを python/blueprints/ から削除するだけで機能が無効化される

エンドポイント:
  POST /mt8821c/execute  名前付きアクションを実行する
  GET  /mt8821c/actions  利用可能なアクション一覧と説明を返す

Request (POST /mt8821c/execute):
  {
    "address": "GPIB0::1::INSTR",
    "action":  "set_dl_power",
    "params":  {"power": -70.0},
    "timeout": 5000
  }

Response:
  {"success": true, "response": "...", "error": "", "address": "...", "command": "..."}
"""
import logging

from flask import Blueprint, current_app, jsonify, request

from instruments.anritsu_mt8821c import AnritsuMT8821C

logger = logging.getLogger(__name__)

blueprint = Blueprint("mt8821c", __name__, url_prefix="/mt8821c")

# アクション定義は AnritsuMT8821C クラスと共有 (単一の定義元)
_ACTIONS = AnritsuMT8821C.ACTIONS


@blueprint.route("/execute", methods=["POST"])
def execute():
    """
    MT8821C の名前付きアクションを実行する

    params が不要なアクション例:
      {"address": "GPIB0::1::INSTR", "action": "identify"}

    params が必要なアクション例:
      {"address": "GPIB0::1::INSTR", "action": "set_dl_power", "params": {"power": -70.0}}
    """
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"success": False, "error": "リクエストボディがJSONではありません", "response": ""}), 400

    address = data.get("address", "").strip()
    action  = data.get("action",  "").strip()
    params  = data.get("params",  {})
    timeout = int(data.get("timeout", 5000))

    if not address or not action:
        return jsonify({"success": False, "error": "address と action は必須です", "response": ""}), 400

    if action not in _ACTIONS:
        return jsonify({
            "success": False,
            "error": f"未知のアクション: '{action}'。GET /mt8821c/actions で一覧を確認してください",
            "response": "",
        }), 400

    cmd_fn, _ = _ACTIONS[action]
    try:
        command = cmd_fn(params)
    except (KeyError, ValueError, TypeError) as e:
        return jsonify({"success": False, "error": f"パラメータエラー: {e}", "response": ""}), 400

    logger.info(
        "→RECV [MT8821C] POST /mt8821c/execute | addr=%s | action=%s | params=%s | scpi=%s | timeout=%dms",
        address, action, params, command, timeout,
    )

    result = current_app.gpib_manager.execute(address, command, timeout)
    logger.info(
        "←SEND [MT8821C] POST /mt8821c/execute | success=%s | resp=%s",
        result["success"], result["response"],
    )
    return jsonify(result), 200 if result["success"] else 500


@blueprint.route("/actions", methods=["GET"])
def list_actions():
    """利用可能なアクション名と説明の一覧を返す"""
    return jsonify({
        "actions": {name: desc for name, (_, desc) in _ACTIONS.items()}
    })
