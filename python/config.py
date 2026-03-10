"""
設定管理・ロギングセットアップ

config/settings.ini を読み込み、アプリケーション全体の設定とロギングを提供する。
CLI方式 (gpib_controller.py) / Flask方式 (server.py) の両方から使用する。
"""
import configparser
import logging
import logging.handlers
import os

# settings.ini のデフォルトパス (このファイルから ../config/settings.ini)
_DEFAULT_INI = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config", "settings.ini")
)


def _load(ini_path: str = None) -> configparser.ConfigParser:
    """INIファイルを読み込む。ファイルが存在しない場合はデフォルト値で動作する。"""
    cfg = configparser.ConfigParser()
    path = os.path.abspath(ini_path or _DEFAULT_INI)
    read_files = cfg.read(path, encoding="utf-8-sig")
    if not read_files:
        logging.getLogger(__name__).warning("設定ファイルが見つかりません: %s (デフォルト値を使用)", path)
    return cfg


def get_server_settings(ini_path: str = None) -> dict:
    """
    [Server] セクションの設定を辞書で返す

    Returns:
        host, port, python_exe, server_script, health_timeout_sec, max_retry
    """
    cfg = _load(ini_path)
    return {
        "host":               cfg.get    ("Server", "Host",               fallback="127.0.0.1"),
        "port":               cfg.getint ("Server", "Port",               fallback=5000),
        "python_exe":         cfg.get    ("Server", "PythonExe",          fallback="python"),
        "server_script":      cfg.get    ("Server", "ServerScript",       fallback=""),
        "health_timeout_sec": cfg.getint ("Server", "HealthTimeoutSec",   fallback=10),
        "max_retry":          cfg.getint ("Server", "MaxRetry",           fallback=1),
    }


def get_lan_settings(ini_path: str = None) -> dict:
    """
    [Lan] セクションの設定を辞書で返す

    Returns:
        default_socket_port, read_termination, write_termination
    """
    def _unescape(s: str) -> str:
        return s.replace("\\n", "\n").replace("\\r", "\r")

    cfg = _load(ini_path)
    return {
        "default_socket_port": cfg.getint("Lan", "DefaultSocketPort", fallback=5025),
        "read_termination":    _unescape(cfg.get("Lan", "ReadTermination",  fallback="\n")),
        "write_termination":   _unescape(cfg.get("Lan", "WriteTermination", fallback="\n")),
    }


def build_visa_address(protocol: str, host: str, port: str = "") -> str:
    """
    接続方式・ホスト・ポートから VISA アドレス文字列を生成する

    Args:
        protocol: "GPIB" / "TCPIP" / "SOCKET" / "HISLIP"
        host: GPIBアドレス番号 or IPアドレス/ホスト名
        port: ポート番号 (SOCKET 時のみ使用。省略時はデフォルトポートを使用)

    Returns:
        VISAアドレス文字列
    """
    p = protocol.upper().strip()
    if p == "GPIB":
        return f"GPIB0::{host}::INSTR"
    if p in ("SOCKET", "TCPIP_SOCKET"):
        actual_port = port.strip() if port.strip() else str(get_lan_settings()["default_socket_port"])
        return f"TCPIP0::{host}::{actual_port}::SOCKET"
    if p in ("HISLIP", "TCPIP_HISLIP"):
        return f"TCPIP0::{host}::hislip0::INSTR"
    if p in ("TCPIP", "VXI11", "LAN", "TCPIP_VXI11"):
        return f"TCPIP0::{host}::INSTR"
    # 不明なプロトコルはホスト値をそのまま返す (フルVISAアドレスが渡された場合)
    return host


def setup_logging(ini_path: str = None) -> None:
    """
    [Logging] セクションに従ってロギングを初期化する。
    コンソールとローテーションファイルの両方に出力する。

    ログディレクトリ: settings.ini の LogDir が相対パスの場合、
                     プロジェクトルート (settings.ini の親ディレクトリの親) を基準に解決する。
    """
    cfg = _load(ini_path)

    level_str  = cfg.get    ("Logging", "Level",       fallback="INFO")
    log_dir    = cfg.get    ("Logging", "LogDir",      fallback="logs")
    filename   = cfg.get    ("Logging", "FileName",    fallback="gpib.log")
    max_bytes  = cfg.getint ("Logging", "MaxBytes",    fallback=1_000_000)
    backup     = cfg.getint ("Logging", "BackupCount", fallback=3)
    fmt        = cfg.get    ("Logging", "Format",
                             fallback="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    level = getattr(logging, level_str.upper(), logging.INFO)

    # 相対パスはプロジェクトルート基準で解決
    if not os.path.isabs(log_dir):
        ini_dir = os.path.dirname(os.path.abspath(ini_path or _DEFAULT_INI))
        log_dir = os.path.normpath(os.path.join(ini_dir, "..", log_dir))

    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, filename)

    handlers = [
        logging.StreamHandler(),
        logging.handlers.RotatingFileHandler(
            log_file, maxBytes=max_bytes, backupCount=backup, encoding="utf-8"
        ),
    ]

    # force=True で既存のハンドラを上書きし、二重登録を防ぐ
    logging.basicConfig(level=level, format=fmt, handlers=handlers, force=True)
    logging.getLogger(__name__).info(
        "ロギング開始 | ファイル: %s | レベル: %s", log_file, level_str
    )
