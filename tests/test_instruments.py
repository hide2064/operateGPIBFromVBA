"""
計測器クラスのテスト

BaseInstrument / LanInstrument / GenericInstrument の動作を検証する。
pyvisa はモックし、実際の GPIB/LAN 接続は行わない。
"""
import pytest

from instruments.base_instrument import BaseInstrument
from instruments.lan_instrument import LanInstrument
from instruments.generic_instrument import GenericInstrument


class TestBaseInstrumentConnectionType:
    """connection_type プロパティのテスト"""

    def test_gpib(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR")
        assert inst.connection_type == "GPIB"

    def test_gpib_various_address(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::22::INSTR")
        assert inst.connection_type == "GPIB"

    def test_tcpip_vxi11(self, mock_pyvisa):
        inst = BaseInstrument("TCPIP0::192.168.1.1::INSTR")
        assert inst.connection_type == "TCPIP_VXI11"

    def test_tcpip_socket(self, mock_pyvisa):
        inst = BaseInstrument("TCPIP0::192.168.1.1::5025::SOCKET")
        assert inst.connection_type == "TCPIP_SOCKET"

    def test_tcpip_hislip(self, mock_pyvisa):
        inst = BaseInstrument("TCPIP0::192.168.1.1::hislip0::INSTR")
        assert inst.connection_type == "TCPIP_HISLIP"

    def test_unknown(self, mock_pyvisa):
        inst = BaseInstrument("USB0::0x1234::INSTR")
        assert inst.connection_type == "UNKNOWN"


class TestBaseInstrumentOpen:
    """open() の動作テスト"""

    def test_socket_sets_read_termination(self, mock_pyvisa):
        inst = BaseInstrument("TCPIP0::192.168.1.1::5025::SOCKET")
        inst.open()
        assert mock_pyvisa.read_termination == "\n"

    def test_socket_sets_write_termination(self, mock_pyvisa):
        inst = BaseInstrument("TCPIP0::192.168.1.1::5025::SOCKET")
        inst.open()
        assert mock_pyvisa.write_termination == "\n"

    def test_gpib_does_not_override_termination(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR")
        inst.open()
        # GPIB は終端文字を明示的に "\n" に設定しない
        assert mock_pyvisa.read_termination != "\n"

    def test_sets_timeout(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR", timeout=3000)
        inst.open()
        assert mock_pyvisa.timeout == 3000

    def test_close_clears_instrument(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR")
        inst.open()
        assert inst._instrument is not None
        inst.close()
        assert inst._instrument is None

    def test_context_manager(self, mock_pyvisa):
        with BaseInstrument("GPIB0::1::INSTR") as inst:
            assert inst._instrument is not None
        assert inst._instrument is None

    def test_write_without_open_raises(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR")
        with pytest.raises(RuntimeError):
            inst.write("*RST")

    def test_query_without_open_raises(self, mock_pyvisa):
        inst = BaseInstrument("GPIB0::1::INSTR")
        with pytest.raises(RuntimeError):
            inst.query("*IDN?")


class TestLanInstrument:
    """LanInstrument ファクトリメソッドのテスト"""

    def test_vxi11_address(self, mock_pyvisa):
        inst = LanInstrument.vxi11("192.168.1.1")
        assert inst.address == "TCPIP0::192.168.1.1::INSTR"
        assert inst.connection_type == "TCPIP_VXI11"

    def test_vxi11_custom_timeout(self, mock_pyvisa):
        inst = LanInstrument.vxi11("192.168.1.1", timeout=3000)
        assert inst._timeout == 3000

    def test_hislip_address(self, mock_pyvisa):
        inst = LanInstrument.hislip("192.168.1.1")
        assert inst.address == "TCPIP0::192.168.1.1::hislip0::INSTR"
        assert inst.connection_type == "TCPIP_HISLIP"

    def test_hislip_custom_index(self, mock_pyvisa):
        inst = LanInstrument.hislip("192.168.1.1", index=1)
        assert inst.address == "TCPIP0::192.168.1.1::hislip1::INSTR"

    def test_socket_address_default_port(self, mock_pyvisa):
        inst = LanInstrument.socket("192.168.1.1")
        assert inst.address == "TCPIP0::192.168.1.1::5025::SOCKET"
        assert inst.connection_type == "TCPIP_SOCKET"

    def test_socket_address_custom_port(self, mock_pyvisa):
        inst = LanInstrument.socket("192.168.1.1", port=5555)
        assert inst.address == "TCPIP0::192.168.1.1::5555::SOCKET"

    def test_socket_sets_termination_on_open(self, mock_pyvisa):
        inst = LanInstrument.socket("192.168.1.1")
        inst.open()
        assert mock_pyvisa.read_termination == "\n"
        assert mock_pyvisa.write_termination == "\n"


class TestGenericInstrumentExecute:
    """GenericInstrument.execute() のテスト"""

    def test_query_command_with_question_mark(self, mock_pyvisa):
        mock_pyvisa.query.return_value = "Keysight,34461A"
        inst = GenericInstrument("GPIB0::1::INSTR")
        inst.open()
        result = inst.execute("*IDN?")
        assert result["success"] is True
        assert result["response"] == "Keysight,34461A"
        mock_pyvisa.query.assert_called_once_with("*IDN?")

    def test_write_command_without_question_mark(self, mock_pyvisa):
        inst = GenericInstrument("GPIB0::1::INSTR")
        inst.open()
        result = inst.execute("VOLT 5.0")
        assert result["success"] is True
        assert result["response"] == "OK"
        mock_pyvisa.write.assert_called_once_with("VOLT 5.0")

    def test_execute_catches_exception(self, mock_pyvisa):
        mock_pyvisa.query.side_effect = Exception("Timeout")
        inst = GenericInstrument("GPIB0::1::INSTR")
        inst.open()
        result = inst.execute("*IDN?")
        assert result["success"] is False
        assert "Timeout" in result["error"]

    def test_query_response_stripped(self, mock_pyvisa):
        mock_pyvisa.query.return_value = "  response with spaces  "
        inst = GenericInstrument("GPIB0::1::INSTR")
        inst.open()
        result = inst.execute("MEAS?")
        assert result["response"] == "response with spaces"
