"""
pytest 共通設定・フィクスチャ

sys.path に python/ ディレクトリを追加し、
pyvisa のモックフィクスチャを提供する。
"""
import os
import sys

import pytest

# python/ ディレクトリをパスに追加
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "python"))


@pytest.fixture
def mock_pyvisa(mocker):
    """
    pyvisa.ResourceManager をモックし、仮想の機器リソースを返す。
    テスト実行時に実際の GPIB/LAN 接続を行わないようにする。
    """
    mock_resource = mocker.MagicMock()
    mock_resource.query.return_value = ""
    mock_resource.timeout = 5000

    mock_rm = mocker.MagicMock()
    mock_rm.open_resource.return_value = mock_resource

    mocker.patch("instruments.base_instrument.pyvisa.ResourceManager", return_value=mock_rm)

    return mock_resource
