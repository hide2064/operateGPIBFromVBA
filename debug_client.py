"""
debug_client.py - GPIB Flask サーバーのデバッグ用CLIクライアント

VBAを使わずにFlaskサーバーのAPIを直接呼び出してテストするツール。

使い方:
  python debug_client.py health
  python debug_client.py execute --address GPIB0::1::INSTR --command "*IDN?"
  python debug_client.py execute --address TCPIP0::192.168.1.10::INSTR --command "*IDN?"
  python debug_client.py execute --address TCPIP0::192.168.1.10::5025::SOCKET --command "*IDN?"
  python debug_client.py connections
  python debug_client.py close --address GPIB0::1::INSTR
  python debug_client.py resources
  python debug_client.py debug
  python debug_client.py log [--lines 50]

オプション:
  --url    サーバーURL (デフォルト: settings.ini の Host/Port または http://127.0.0.1:5000)
  --config settings.ini のパス
"""
import argparse
import json
import os
import sys
import urllib.error
import urllib.request


# -----------------------------------------------------------------------
# settings.ini 読み込み (configparser 不要ならフォールバック)
# -----------------------------------------------------------------------

def _read_ini(config_path: str = None) -> dict:
    """settings.ini から最低限の設定を読む。失敗したらデフォルト値。"""
    defaults = {"host": "127.0.0.1", "port": 5000, "log_dir": "logs", "log_file": "gpib.log"}
    try:
        import configparser
        ini = os.path.normpath(
            config_path or os.path.join(os.path.dirname(__file__), "config", "settings.ini")
        )
        cfg = configparser.ConfigParser()
        cfg.read(ini, encoding="utf-8-sig")
        defaults["host"]     = cfg.get    ("Server",  "Host",      fallback=defaults["host"])
        defaults["port"]     = cfg.getint ("Server",  "Port",      fallback=defaults["port"])
        defaults["log_dir"]  = cfg.get    ("Logging", "LogDir",    fallback=defaults["log_dir"])
        defaults["log_file"] = cfg.get    ("Logging", "FileName",  fallback=defaults["log_file"])
        # 相対パス → プロジェクトルート基準
        if not os.path.isabs(defaults["log_dir"]):
            ini_dir = os.path.dirname(os.path.abspath(ini))
            defaults["log_dir"] = os.path.normpath(os.path.join(ini_dir, "..", defaults["log_dir"]))
    except Exception:
        pass
    return defaults


# -----------------------------------------------------------------------
# HTTP ヘルパー
# -----------------------------------------------------------------------

def _get(url: str) -> dict:
    req = urllib.request.Request(url, method="GET")
    with urllib.request.urlopen(req, timeout=10) as res:
        return json.loads(res.read().decode())


def _post(url: str, body: dict) -> dict:
    data = json.dumps(body, ensure_ascii=False).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            return json.loads(res.read().decode())
    except urllib.error.HTTPError as e:
        return json.loads(e.read().decode())


def _delete(url: str) -> dict:
    req = urllib.request.Request(url, method="DELETE")
    try:
        with urllib.request.urlopen(req, timeout=10) as res:
            return json.loads(res.read().decode())
    except urllib.error.HTTPError as e:
        return json.loads(e.read().decode())


# -----------------------------------------------------------------------
# 出力ヘルパー
# -----------------------------------------------------------------------

def _print(label: str, data: dict):
    print(f"\n[{label}]")
    print(json.dumps(data, ensure_ascii=False, indent=2))


def _ok(msg: str):
    print(f"  OK  {msg}")


def _err(msg: str):
    print(f" ERR  {msg}", file=sys.stderr)


# -----------------------------------------------------------------------
# コマンド実装
# -----------------------------------------------------------------------

def cmd_health(base_url: str, **_):
    try:
        res = _get(f"{base_url}/health")
        _print("health", res)
        if res.get("status") == "ok":
            _ok("サーバーは正常稼働中")
    except Exception as e:
        _err(f"サーバーに接続できません: {e}")
        _err(f"  → サーバーが起動しているか確認: start_server.bat")


def cmd_execute(base_url: str, address: str, command: str, timeout: int = 5000, **_):
    if not address or not command:
        _err("--address と --command は必須です")
        return
    print(f"\n送信: address={address}  command={command}  timeout={timeout}ms")
    res = _post(f"{base_url}/execute", {"address": address, "command": command, "timeout": timeout})
    _print("execute", res)
    if res.get("success"):
        _ok(f"応答: {res.get('response', '(応答なし)')}")
    else:
        _err(f"失敗: {res.get('error')}")


def cmd_connections(base_url: str, **_):
    res = _get(f"{base_url}/connections")
    _print("connections", res)
    conns = res.get("connections", [])
    if conns:
        print(f"  接続中: {len(conns)} 件")
        for c in conns:
            print(f"    - {c}")
    else:
        print("  接続中のリソースはありません")


def cmd_close(base_url: str, address: str, **_):
    if not address:
        _err("--address が必要です")
        return
    res = _delete(f"{base_url}/connections/{address}")
    _print("close", res)


def cmd_resources(base_url: str, **_):
    res = _get(f"{base_url}/resources")
    _print("resources", res)
    resources = res.get("resources", [])
    if resources:
        print(f"  VISAリソース: {len(resources)} 件")
        for r in resources:
            print(f"    - {r}")
    else:
        print("  VISAリソースが見つかりません (GPIB/LANカードが接続されていないか、機器がオフ)")


def cmd_debug(base_url: str, **_):
    try:
        res = _get(f"{base_url}/debug")
        _print("debug", res)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            _err("/debug エンドポイントが存在しません。server.py が古い可能性があります。")
        else:
            _err(f"HTTP {e.code}: {e.reason}")
    except Exception as e:
        _err(f"エラー: {e}")


def cmd_log(ini_settings: dict, lines: int = 30, **_):
    log_path = os.path.join(ini_settings["log_dir"], ini_settings["log_file"])
    if not os.path.exists(log_path):
        _err(f"ログファイルが見つかりません: {log_path}")
        _err("  → サーバーをまだ起動したことがないか、LogDir 設定を確認してください")
        return
    print(f"\n[log] {log_path} (最新 {lines} 行)")
    print("-" * 60)
    with open(log_path, encoding="utf-8", errors="replace") as f:
        all_lines = f.readlines()
    for line in all_lines[-lines:]:
        print(line, end="")
    print("-" * 60)


# -----------------------------------------------------------------------
# エントリポイント
# -----------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="GPIB Flask サーバーのデバッグCLIクライアント",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--url",     default=None, help="サーバーURL (例: http://127.0.0.1:5000)")
    parser.add_argument("--config",  default=None, help="settings.ini のパス")
    parser.add_argument("--address", default=None, help="VISAアドレス (execute/close コマンド用)")
    parser.add_argument("--command", default=None, help="SCPIコマンド (execute コマンド用)")
    parser.add_argument("--timeout", default=5000, type=int, help="タイムアウト ms (default: 5000)")
    parser.add_argument("--lines",   default=30,   type=int, help="ログ表示行数 (default: 30)")
    parser.add_argument(
        "cmd",
        choices=["health", "execute", "connections", "close", "resources", "debug", "log"],
        help="実行するコマンド",
    )

    args = parser.parse_args()
    ini_settings = _read_ini(args.config)

    base_url = args.url or f"http://{ini_settings['host']}:{ini_settings['port']}"

    dispatch = {
        "health":      cmd_health,
        "execute":     cmd_execute,
        "connections": cmd_connections,
        "close":       cmd_close,
        "resources":   cmd_resources,
        "debug":       cmd_debug,
        "log":         cmd_log,
    }
    dispatch[args.cmd](
        base_url=base_url,
        address=args.address,
        command=args.command,
        timeout=args.timeout,
        lines=args.lines,
        ini_settings=ini_settings,
    )


if __name__ == "__main__":
    main()
