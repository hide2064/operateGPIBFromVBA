"""
config.py のテスト

build_visa_address() / get_lan_settings() / get_server_settings() を検証する。
"""
import pytest

from config import build_visa_address, get_lan_settings, get_server_settings


class TestBuildVisaAddress:
    """build_visa_address() のテスト"""

    def test_gpib(self):
        assert build_visa_address("GPIB", "1") == "GPIB0::1::INSTR"
        assert build_visa_address("GPIB", "22") == "GPIB0::22::INSTR"

    def test_tcpip_vxi11_variants(self):
        for proto in ("TCPIP", "VXI11", "LAN", "TCPIP_VXI11"):
            result = build_visa_address(proto, "192.168.1.1")
            assert result == "TCPIP0::192.168.1.1::INSTR", f"プロトコル '{proto}' の結果が不正: {result}"

    def test_socket_default_port(self):
        result = build_visa_address("SOCKET", "192.168.1.1")
        assert result == "TCPIP0::192.168.1.1::5025::SOCKET"

    def test_socket_custom_port(self):
        result = build_visa_address("SOCKET", "192.168.1.1", port="5555")
        assert result == "TCPIP0::192.168.1.1::5555::SOCKET"

    def test_socket_alias(self):
        result = build_visa_address("TCPIP_SOCKET", "192.168.1.1", port="5025")
        assert result == "TCPIP0::192.168.1.1::5025::SOCKET"

    def test_hislip(self):
        for proto in ("HISLIP", "TCPIP_HISLIP"):
            result = build_visa_address(proto, "192.168.1.1")
            assert result == "TCPIP0::192.168.1.1::hislip0::INSTR"

    def test_unknown_protocol_returns_host(self):
        # 不明なプロトコルは host をそのまま VISA アドレスとして返す
        full_address = "GPIB0::1::INSTR"
        assert build_visa_address("UNKNOWN", full_address) == full_address

    def test_case_insensitive(self):
        assert build_visa_address("gpib", "1") == build_visa_address("GPIB", "1")
        assert build_visa_address("tcpip", "192.168.1.1") == build_visa_address("TCPIP", "192.168.1.1")
        assert build_visa_address("socket", "192.168.1.1") == build_visa_address("SOCKET", "192.168.1.1")


class TestGetLanSettings:
    """get_lan_settings() のテスト"""

    def test_default_socket_port(self):
        settings = get_lan_settings()
        assert settings["default_socket_port"] == 5025

    def test_default_termination(self):
        settings = get_lan_settings()
        assert settings["read_termination"] == "\n"
        assert settings["write_termination"] == "\n"

    def test_returns_all_keys(self):
        settings = get_lan_settings()
        assert "default_socket_port" in settings
        assert "read_termination" in settings
        assert "write_termination" in settings


class TestGetServerSettings:
    """get_server_settings() のテスト"""

    def test_default_host(self):
        settings = get_server_settings()
        assert settings["host"] == "127.0.0.1"

    def test_default_port(self):
        settings = get_server_settings()
        assert settings["port"] == 5000

    def test_default_max_retry(self):
        settings = get_server_settings()
        assert settings["max_retry"] == 1

    def test_returns_all_keys(self):
        settings = get_server_settings()
        for key in ("host", "port", "python_exe", "server_script", "health_timeout_sec", "max_retry"):
            assert key in settings, f"キー '{key}' が返り値に含まれていません"
