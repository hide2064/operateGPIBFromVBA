"""
Anritsu MT8821C クラスのテスト

ACTIONS 辞書の定義とコマンド生成、各メソッドの動作を検証する。
"""
import pytest

from instruments.anritsu_mt8821c import AnritsuMT8821C


EXPECTED_ACTIONS = [
    "identify", "reset", "preset", "get_error",
    "get_dl_power", "set_dl_power",
    "get_band", "set_band",
    "get_channel", "set_channel",
    "call_connect", "call_disconnect", "get_call_status",
    "measure_ul_power",
]


class TestMT8821CActionsDict:
    """ACTIONS クラス変数のテスト"""

    def test_all_expected_actions_exist(self):
        for action in EXPECTED_ACTIONS:
            assert action in AnritsuMT8821C.ACTIONS, f"アクション '{action}' が未定義"

    def test_each_action_has_callable_and_description(self):
        for name, value in AnritsuMT8821C.ACTIONS.items():
            assert len(value) == 2, f"ACTIONS['{name}'] は (callable, str) のタプルであること"
            cmd_fn, desc = value
            assert callable(cmd_fn), f"ACTIONS['{name}'][0] が callable ではありません"
            assert isinstance(desc, str) and desc, f"ACTIONS['{name}'][1] が空文字列です"

    # --- システムコマンド ---
    def test_identify_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["identify"]
        assert cmd_fn({}) == "*IDN?"

    def test_reset_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["reset"]
        assert cmd_fn({}) == "*RST"

    def test_preset_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["preset"]
        assert cmd_fn({}) == "SYSTem:PRESet"

    def test_get_error_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["get_error"]
        assert cmd_fn({}) == "SYSTem:ERRor?"

    # --- セル設定コマンド ---
    def test_get_dl_power_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["get_dl_power"]
        assert cmd_fn({}) == "BS:OLVL?"

    def test_set_dl_power_command_negative(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_dl_power"]
        assert cmd_fn({"power": -70.0}) == "BS:OLVL -70.0"

    def test_set_dl_power_command_decimal(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_dl_power"]
        assert cmd_fn({"power": -50.5}) == "BS:OLVL -50.5"

    def test_set_dl_power_missing_param_raises(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_dl_power"]
        with pytest.raises(KeyError):
            cmd_fn({})

    def test_get_band_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["get_band"]
        assert cmd_fn({}) == "BAND?"

    def test_set_band_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_band"]
        assert cmd_fn({"band": 1}) == "BAND 1"
        assert cmd_fn({"band": 28}) == "BAND 28"

    def test_set_band_missing_param_raises(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_band"]
        with pytest.raises(KeyError):
            cmd_fn({})

    def test_get_channel_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["get_channel"]
        assert cmd_fn({}) == "CHANL?"

    def test_set_channel_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["set_channel"]
        assert cmd_fn({"channel": 300}) == "CHANL 300"

    # --- コール処理コマンド ---
    def test_call_connect_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["call_connect"]
        assert cmd_fn({}) == "CALLSO"

    def test_call_disconnect_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["call_disconnect"]
        assert cmd_fn({}) == "CALLEND"

    def test_get_call_status_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["get_call_status"]
        assert cmd_fn({}) == "CALLSTAT?"

    # --- 測定コマンド ---
    def test_measure_ul_power_command(self):
        cmd_fn, _ = AnritsuMT8821C.ACTIONS["measure_ul_power"]
        assert cmd_fn({}) == "MEAS:UL:POW?"


class TestMT8821CMethods:
    """AnritsuMT8821C インスタンスメソッドのテスト"""

    def test_identify(self, mock_pyvisa):
        mock_pyvisa.query.return_value = "Anritsu,MT8821C,001,1.0"
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.identify()
        assert result["success"] is True
        assert result["response"] == "Anritsu,MT8821C,001,1.0"
        mock_pyvisa.query.assert_called_with("*IDN?")

    def test_preset(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.preset()
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("SYSTem:PRESet")

    def test_set_dl_power(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.set_dl_power(-70.0)
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("BS:OLVL -70.0")

    def test_set_band(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.set_band(1)
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("BAND 1")

    def test_set_channel(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.set_channel(300)
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("CHANL 300")

    def test_call_connect(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.call_connect()
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("CALLSO")

    def test_call_disconnect(self, mock_pyvisa):
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.call_disconnect()
        assert result["success"] is True
        mock_pyvisa.write.assert_called_with("CALLEND")

    def test_get_call_status(self, mock_pyvisa):
        mock_pyvisa.query.return_value = "IDLE"
        inst = AnritsuMT8821C("GPIB0::1::INSTR")
        inst.open()
        result = inst.get_call_status()
        assert result["success"] is True
        assert result["response"] == "IDLE"

    def test_lan_connection(self, mock_pyvisa):
        """LAN接続でも同じコマンドが動作することを確認"""
        mock_pyvisa.query.return_value = "Anritsu,MT8821C"
        inst = AnritsuMT8821C("TCPIP0::192.168.1.100::INSTR")
        inst.open()
        result = inst.identify()
        assert result["success"] is True
        assert inst.connection_type == "TCPIP_VXI11"
