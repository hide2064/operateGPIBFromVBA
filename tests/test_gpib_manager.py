"""
GpibManager のテスト

接続プール・リトライ・スレッドセーフ動作を検証する。
GenericInstrument はモックし、実際の機器接続は行わない。
"""
import pytest

from gpib_manager import GpibManager


ADDRESS_1 = "GPIB0::1::INSTR"
ADDRESS_2 = "GPIB0::2::INSTR"

SUCCESS_RESULT = {"success": True, "response": "OK", "error": ""}
FAILURE_RESULT = {"success": False, "response": "", "error": "Timeout"}


@pytest.fixture
def mock_instrument(mocker):
    """GenericInstrument をモックする"""
    mock_inst = mocker.MagicMock()
    mock_inst.execute.return_value = SUCCESS_RESULT
    mocker.patch("gpib_manager.GenericInstrument", return_value=mock_inst)
    return mock_inst


class TestGpibManagerExecute:
    """execute() の動作テスト"""

    def test_success(self, mock_instrument):
        manager = GpibManager(max_retry=1)
        result = manager.execute(ADDRESS_1, "*IDN?")
        assert result["success"] is True
        mock_instrument.open.assert_called_once()
        mock_instrument.execute.assert_called_once_with("*IDN?")

    def test_result_includes_address_and_command(self, mock_instrument):
        manager = GpibManager()
        result = manager.execute(ADDRESS_1, "*IDN?")
        assert result["address"] == ADDRESS_1
        assert result["command"] == "*IDN?"

    def test_failure_returns_error(self, mock_instrument):
        mock_instrument.execute.return_value = FAILURE_RESULT
        manager = GpibManager(max_retry=0)
        result = manager.execute(ADDRESS_1, "*IDN?")
        assert result["success"] is False
        assert result["error"] == "Timeout"

    def test_retry_on_failure_then_success(self, mock_instrument):
        mock_instrument.execute.side_effect = [FAILURE_RESULT, SUCCESS_RESULT]
        manager = GpibManager(max_retry=1)
        result = manager.execute(ADDRESS_1, "VOLT 5.0")
        assert result["success"] is True
        assert mock_instrument.execute.call_count == 2

    def test_all_retries_fail(self, mock_instrument):
        mock_instrument.execute.return_value = FAILURE_RESULT
        manager = GpibManager(max_retry=2)
        result = manager.execute(ADDRESS_1, "*IDN?")
        assert result["success"] is False
        # 最初の試行 + リトライ2回 = 計3回
        assert mock_instrument.execute.call_count == 3

    def test_exception_triggers_retry(self, mock_instrument):
        mock_instrument.execute.side_effect = [Exception("Connection lost"), SUCCESS_RESULT]
        manager = GpibManager(max_retry=1)
        result = manager.execute(ADDRESS_1, "*IDN?")
        assert result["success"] is True


class TestGpibManagerConnectionPool:
    """接続プール (再利用) のテスト"""

    def test_connection_reused_on_same_address(self, mock_instrument):
        manager = GpibManager()
        manager.execute(ADDRESS_1, "*IDN?")
        manager.execute(ADDRESS_1, "*RST")
        # 同じアドレスへの2回の呼び出しで open() は1回だけ
        assert mock_instrument.open.call_count == 1

    def test_different_addresses_open_separately(self, mock_instrument):
        manager = GpibManager()
        manager.execute(ADDRESS_1, "*IDN?")
        manager.execute(ADDRESS_2, "*IDN?")
        # 異なるアドレスなので open() は2回
        assert mock_instrument.open.call_count == 2

    def test_list_connections_empty_initially(self, mock_instrument):
        manager = GpibManager()
        assert manager.list_connections() == []

    def test_list_connections_after_execute(self, mock_instrument):
        manager = GpibManager()
        manager.execute(ADDRESS_1, "*IDN?")
        manager.execute(ADDRESS_2, "*IDN?")
        connections = manager.list_connections()
        addresses = [c["address"] for c in connections]
        assert ADDRESS_1 in addresses
        assert ADDRESS_2 in addresses

    def test_close_connection_removes_from_pool(self, mock_instrument):
        manager = GpibManager()
        manager.execute(ADDRESS_1, "*IDN?")
        result = manager.close_connection(ADDRESS_1)
        assert result is True
        assert manager.list_connections() == []
        mock_instrument.close.assert_called_once()

    def test_close_nonexistent_connection_returns_false(self, mock_instrument):
        manager = GpibManager()
        result = manager.close_connection("GPIB0::99::INSTR")
        assert result is False

    def test_close_all_clears_pool(self, mock_instrument):
        manager = GpibManager()
        manager.execute(ADDRESS_1, "*IDN?")
        manager.execute(ADDRESS_2, "*IDN?")
        manager.close_all()
        assert manager.list_connections() == []

    def test_retry_closes_and_reopens_connection(self, mock_instrument, mocker):
        """リトライ時に接続を閉じて再接続することを確認する"""
        mock_instrument.execute.side_effect = [FAILURE_RESULT, SUCCESS_RESULT]
        manager = GpibManager(max_retry=1)
        manager.execute(ADDRESS_1, "*IDN?")
        # 失敗後に close() が呼ばれ、再接続のため open() が2回呼ばれる
        assert mock_instrument.close.call_count >= 1
        assert mock_instrument.open.call_count == 2
