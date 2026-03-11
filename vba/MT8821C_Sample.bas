Attribute VB_Name = "MT8821C_Sample"
'=============================================================================
' MT8821C_Sample.bas - MT8821C 動作確認サンプル集
'
' GPIB 接続・LAN 接続それぞれで MT8821C の各操作を試すためのサンプル。
' Config/Control シートは使わず、アドレスをコード内に直書きして単体で動く。
'
' 【前提】
'   - AppConfig.bas, GpibMT8821C.bas がインポートされていること
'   - Flaskサーバーが起動済みであること (start_server.bat を実行)
'
' 【使い方】
'   1. 下記の定数セクションをご自身の環境に合わせて変更する
'   2. VBAエディタのマクロ実行 (F5) またはボタンに割り当てて実行する
'
' 【サンプル一覧】
'   [GPIB]
'     Sample_GPIB_BasicCheck   - 識別・リセット・エラー確認
'     Sample_GPIB_LteSetup     - Band / Channel / DL Power 設定
'     Sample_GPIB_CallTest     - コール接続 -> UL電力測定 -> 切断
'     RunAllSamples_GPIB       - 上記3つを順番に実行
'
'   [LAN - VXI-11]
'     Sample_LAN_BasicCheck    - 識別・リセット・エラー確認 (VXI-11)
'     Sample_LAN_LteSetup      - Band / Channel / DL Power 設定 (VXI-11)
'     Sample_LAN_CallTest      - コール接続 -> UL電力測定 -> 切断 (VXI-11)
'     RunAllSamples_LAN        - 上記3つを順番に実行
'
'   [LAN - Raw Socket]
'     Sample_Socket_BasicCheck - 識別・リセット・エラー確認 (Raw Socket)
'
'   [LAN - HiSLIP]
'     Sample_HiSLIP_BasicCheck - 識別・リセット・エラー確認 (HiSLIP)
'=============================================================================
Option Explicit

'=============================================================================
' ★ ここを環境に合わせて変更する ★
'=============================================================================

' GPIB アドレス (GPIBボードアドレス::機器アドレス)
Private Const ADDR_GPIB     As String = "GPIB0::1::INSTR"

' LAN - VXI-11 (標準的な LAN 接続)
Private Const ADDR_LAN_VXI  As String = "TCPIP0::192.168.1.10::INSTR"

' LAN - Raw Socket (ポート 5025 が一般的)
Private Const ADDR_LAN_SOCK As String = "TCPIP0::192.168.1.10::5025::SOCKET"

' LAN - HiSLIP (高速 LAN インタフェース)
Private Const ADDR_LAN_HISL As String = "TCPIP0::192.168.1.10::hislip0::INSTR"

' LTE テスト設定値
Private Const LTE_BAND     As Integer = 1      ' Band 1 (2GHz)
Private Const LTE_CHANNEL  As Long    = 300    ' チャネル番号
Private Const LTE_DL_POWER As Double  = -70.0  ' DL出力レベル (dBm)

'=============================================================================
' [GPIB] 基本動作確認
' 実行内容: 識別 -> リセット -> プリセット -> エラー確認
'=============================================================================
Public Sub Sample_GPIB_BasicCheck()
    Const TITLE As String = "[GPIB] 基本動作確認"
    Dim addr As String: addr = ADDR_GPIB
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' 識別
    Dim idn As String
    idn = MT8821C_Identify(addr)
    log = log & "*IDN?  : " & idn & vbCrLf

    ' リセット
    Dim rst As String
    rst = MT8821C_Reset(addr)
    log = log & "*RST   : " & IIf(rst = "", "OK", rst) & vbCrLf

    ' プリセット
    Dim preset As String
    preset = MT8821C_Preset(addr)
    log = log & "PRESET : " & IIf(preset = "", "OK", preset) & vbCrLf

    ' エラー確認
    Dim err As String
    err = MT8821C_GetError(addr)
    log = log & "ERROR? : " & err & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [GPIB] LTE 設定サンプル
' 実行内容: Band設定 -> Channel設定 -> DL Power設定 -> 各設定値を読み返す
'=============================================================================
Public Sub Sample_GPIB_LteSetup()
    Const TITLE As String = "[GPIB] LTE 設定"
    Dim addr As String: addr = ADDR_GPIB
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' Band 設定
    Call MT8821C_SetBand(addr, LTE_BAND)
    Dim band As String
    band = MT8821C_GetBand(addr)
    log = log & "Band   : Set=" & LTE_BAND & "  Get=" & band & vbCrLf

    ' Channel 設定
    Call MT8821C_SetChannel(addr, LTE_CHANNEL)
    Dim ch As String
    ch = MT8821C_GetChannel(addr)
    log = log & "Channel: Set=" & LTE_CHANNEL & "  Get=" & ch & vbCrLf

    ' DL Power 設定
    Call MT8821C_SetDlPower(addr, LTE_DL_POWER)
    Dim pwr As String
    pwr = MT8821C_GetDlPower(addr)
    log = log & "DL Pwr : Set=" & LTE_DL_POWER & "dBm  Get=" & pwr & vbCrLf

    ' エラー確認
    log = log & vbCrLf & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [GPIB] コール接続テスト
' 実行内容: コール接続 -> 接続状態確認 -> UL電力測定 -> コール切断
'=============================================================================
Public Sub Sample_GPIB_CallTest()
    Const TITLE As String = "[GPIB] コール接続テスト"
    Dim addr As String: addr = ADDR_GPIB
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    If MsgBox("UE (端末) を接続してからOKを押してください。" & vbCrLf & _
              "接続先: " & addr, vbOKCancel + vbQuestion, TITLE) = vbCancel Then
        Exit Sub
    End If

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' コール接続開始
    Application.StatusBar = "コール接続中..."
    Call MT8821C_CallConnect(addr)
    log = log & "CALLSO (接続要求): 送信済み" & vbCrLf

    ' 接続状態確認 (3秒待って確認)
    Application.Wait Now + TimeSerial(0, 0, 3)
    Dim stat As String
    stat = MT8821C_GetCallStatus(addr)
    log = log & "CALLSTAT?        : " & stat & vbCrLf

    ' UL 電力測定
    Application.StatusBar = "UL電力測定中..."
    Dim ulpwr As String
    ulpwr = MT8821C_MeasureUlPower(addr)
    log = log & "UL Power (dBm)   : " & ulpwr & vbCrLf

    ' コール切断
    Application.StatusBar = "コール切断中..."
    Call MT8821C_CallDisconnect(addr)
    log = log & "CALLEND (切断)   : 送信済み" & vbCrLf

    ' エラー確認
    log = log & vbCrLf & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [GPIB] 全サンプルを順番に実行
'=============================================================================
Public Sub RunAllSamples_GPIB()
    Call Sample_GPIB_BasicCheck
    Call Sample_GPIB_LteSetup
    Call Sample_GPIB_CallTest
End Sub

'=============================================================================
' [LAN - VXI-11] 基本動作確認
' 実行内容: 識別 -> リセット -> プリセット -> エラー確認
'=============================================================================
Public Sub Sample_LAN_BasicCheck()
    Const TITLE As String = "[LAN VXI-11] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_VXI
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String
    idn = MT8821C_Identify(addr)
    log = log & "*IDN?  : " & idn & vbCrLf

    Dim rst As String
    rst = MT8821C_Reset(addr)
    log = log & "*RST   : " & IIf(rst = "", "OK", rst) & vbCrLf

    Dim preset As String
    preset = MT8821C_Preset(addr)
    log = log & "PRESET : " & IIf(preset = "", "OK", preset) & vbCrLf

    log = log & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [LAN - VXI-11] LTE 設定サンプル
'=============================================================================
Public Sub Sample_LAN_LteSetup()
    Const TITLE As String = "[LAN VXI-11] LTE 設定"
    Dim addr As String: addr = ADDR_LAN_VXI
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Call MT8821C_SetBand(addr, LTE_BAND)
    log = log & "Band   : Set=" & LTE_BAND & "  Get=" & MT8821C_GetBand(addr) & vbCrLf

    Call MT8821C_SetChannel(addr, LTE_CHANNEL)
    log = log & "Channel: Set=" & LTE_CHANNEL & "  Get=" & MT8821C_GetChannel(addr) & vbCrLf

    Call MT8821C_SetDlPower(addr, LTE_DL_POWER)
    log = log & "DL Pwr : Set=" & LTE_DL_POWER & "dBm  Get=" & MT8821C_GetDlPower(addr) & vbCrLf

    log = log & vbCrLf & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [LAN - VXI-11] コール接続テスト
'=============================================================================
Public Sub Sample_LAN_CallTest()
    Const TITLE As String = "[LAN VXI-11] コール接続テスト"
    Dim addr As String: addr = ADDR_LAN_VXI
    Dim log As String

    log = "接続先: " & addr & vbCrLf & vbCrLf

    If MsgBox("UE (端末) を接続してからOKを押してください。" & vbCrLf & _
              "接続先: " & addr, vbOKCancel + vbQuestion, TITLE) = vbCancel Then
        Exit Sub
    End If

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Call MT8821C_CallConnect(addr)
    log = log & "CALLSO (接続要求): 送信済み" & vbCrLf

    Application.Wait Now + TimeSerial(0, 0, 3)
    log = log & "CALLSTAT?        : " & MT8821C_GetCallStatus(addr) & vbCrLf

    log = log & "UL Power (dBm)   : " & MT8821C_MeasureUlPower(addr) & vbCrLf

    Call MT8821C_CallDisconnect(addr)
    log = log & "CALLEND (切断)   : 送信済み" & vbCrLf

    log = log & vbCrLf & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [LAN - VXI-11] 全サンプルを順番に実行
'=============================================================================
Public Sub RunAllSamples_LAN()
    Call Sample_LAN_BasicCheck
    Call Sample_LAN_LteSetup
    Call Sample_LAN_CallTest
End Sub

'=============================================================================
' [LAN - Raw Socket] 基本動作確認
' ポート5025に Raw TCP で接続する。終端文字は LF (\n) を使用する。
' MT8821C が Raw Socket に対応している場合に使用する。
'=============================================================================
Public Sub Sample_Socket_BasicCheck()
    Const TITLE As String = "[LAN Socket] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_SOCK
    Dim log As String

    log = "接続先: " & addr & vbCrLf & _
          "(Raw Socket - ポート5025, 終端文字 LF)" & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String
    idn = MT8821C_Identify(addr)
    log = log & "*IDN?  : " & idn & vbCrLf

    Dim rst As String
    rst = MT8821C_Reset(addr)
    log = log & "*RST   : " & IIf(rst = "", "OK", rst) & vbCrLf

    log = log & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "※ MT8821C が Raw Socket 未対応の場合は VXI-11 (TCPIP::host::INSTR) を使用してください。", _
           vbCritical, TITLE
End Sub

'=============================================================================
' [LAN - HiSLIP] 基本動作確認
' HiSLIP (High-Speed LAN Instrument Protocol) で接続する。
' MT8821C が HiSLIP に対応している場合に使用する (高速・安定)。
'=============================================================================
Public Sub Sample_HiSLIP_BasicCheck()
    Const TITLE As String = "[LAN HiSLIP] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_HISL
    Dim log As String

    log = "接続先: " & addr & vbCrLf & _
          "(HiSLIP - hislip0)" & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String
    idn = MT8821C_Identify(addr)
    log = log & "*IDN?  : " & idn & vbCrLf

    Dim rst As String
    rst = MT8821C_Reset(addr)
    log = log & "*RST   : " & IIf(rst = "", "OK", rst) & vbCrLf

    log = log & "ERROR? : " & MT8821C_GetError(addr) & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "※ MT8821C が HiSLIP 未対応の場合は VXI-11 (TCPIP::host::INSTR) を使用してください。", _
           vbCritical, TITLE
End Sub

'=============================================================================
' 接続方式の比較確認
' GPIB と LAN(VXI-11) の両方で *IDN? を実行して接続を確認する
'=============================================================================
Public Sub Sample_ConnectionComparison()
    Const TITLE As String = "接続方式 比較確認"
    Dim log As String

    log = "各接続方式での *IDN? 結果" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

    Application.StatusBar = "GPIB 確認中..."
    On Error Resume Next
    Dim gpib_res As String
    gpib_res = MT8821C_Identify(ADDR_GPIB)
    If Err.Number <> 0 Then gpib_res = "ERROR: " & Err.Description
    Err.Clear
    log = log & "[GPIB]" & vbCrLf & "  addr: " & ADDR_GPIB & vbCrLf & "  IDN : " & gpib_res & vbCrLf & vbCrLf

    Application.StatusBar = "LAN (VXI-11) 確認中..."
    Dim vxi_res As String
    vxi_res = MT8821C_Identify(ADDR_LAN_VXI)
    If Err.Number <> 0 Then vxi_res = "ERROR: " & Err.Description
    Err.Clear
    log = log & "[LAN VXI-11]" & vbCrLf & "  addr: " & ADDR_LAN_VXI & vbCrLf & "  IDN : " & vxi_res & vbCrLf & vbCrLf

    Application.StatusBar = "LAN (Socket) 確認中..."
    Dim sock_res As String
    sock_res = MT8821C_Identify(ADDR_LAN_SOCK)
    If Err.Number <> 0 Then sock_res = "ERROR: " & Err.Description
    Err.Clear
    log = log & "[LAN Socket]" & vbCrLf & "  addr: " & ADDR_LAN_SOCK & vbCrLf & "  IDN : " & sock_res & vbCrLf & vbCrLf

    Application.StatusBar = "LAN (HiSLIP) 確認中..."
    Dim hisl_res As String
    hisl_res = MT8821C_Identify(ADDR_LAN_HISL)
    If Err.Number <> 0 Then hisl_res = "ERROR: " & Err.Description
    Err.Clear
    On Error GoTo 0

    log = log & "[LAN HiSLIP]" & vbCrLf & "  addr: " & ADDR_LAN_HISL & vbCrLf & "  IDN : " & hisl_res & vbCrLf

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
End Sub
