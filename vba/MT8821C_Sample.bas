Attribute VB_Name = "MT8821C_Sample"
'=============================================================================
' MT8821C_Sample.bas - MT8821C 動作確認サンプル集
'
' 各サンプルはコマンド実行後に Result シートへ結果を1行ずつ自動記録する。
' ResultSheet.bas がインポートされていることが前提。
'
' 【前提モジュール】
'   AppConfig.bas / GpibMT8821C.bas / ResultSheet.bas
'
' 【使い方】
'   1. 下記の定数セクションを環境に合わせて変更する
'   2. 各マクロを VBAエディタから直接実行 (F5) するか、ボタンに割り当てる
'
' 【サンプル一覧】
'   [GPIB]
'     Sample_GPIB_BasicCheck    : 識別・リセット・エラー確認
'     Sample_GPIB_LteSetup      : Band / Channel / DL Power 設定
'     Sample_GPIB_CallTest      : コール接続 -> UL電力測定 -> 切断
'     RunAllSamples_GPIB        : 上記3つを連続実行
'
'   [LAN - VXI-11]
'     Sample_LAN_BasicCheck     : 識別・リセット・エラー確認
'     Sample_LAN_LteSetup       : Band / Channel / DL Power 設定
'     Sample_LAN_CallTest       : コール接続 -> UL電力測定 -> 切断
'     RunAllSamples_LAN         : 上記3つを連続実行
'
'   [LAN - Raw Socket / HiSLIP]
'     Sample_Socket_BasicCheck  : 識別・リセット・エラー確認
'     Sample_HiSLIP_BasicCheck  : 識別・リセット・エラー確認
'
'   [比較]
'     Sample_ConnectionComparison : 全接続方式で *IDN? を一括比較
'=============================================================================
Option Explicit

'=============================================================================
' ★ ここを環境に合わせて変更する ★
'=============================================================================

' GPIB アドレス
Private Const ADDR_GPIB     As String = "GPIB0::1::INSTR"

' LAN - VXI-11
Private Const ADDR_LAN_VXI  As String = "TCPIP0::192.168.1.10::INSTR"

' LAN - Raw Socket
Private Const ADDR_LAN_SOCK As String = "TCPIP0::192.168.1.10::5025::SOCKET"

' LAN - HiSLIP
Private Const ADDR_LAN_HISL As String = "TCPIP0::192.168.1.10::hislip0::INSTR"

' LTE テスト設定値
Private Const LTE_BAND     As Integer = 1      ' Band 1 (2GHz)
Private Const LTE_CHANNEL  As Long    = 300    ' チャネル番号
Private Const LTE_DL_POWER As Double  = -70.0  ' DL出力レベル (dBm)

' 機器名 (Result シートに記録される)
Private Const DEVICE_NAME   As String = "MT8821C"

'=============================================================================
' [GPIB] 基本動作確認
' 実行内容: 識別 -> リセット -> プリセット -> エラー確認
'=============================================================================
Public Sub Sample_GPIB_BasicCheck()
    Const TITLE As String = "[GPIB] 基本動作確認"
    Dim addr As String: addr = ADDR_GPIB
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' 識別
    Dim idn As String: idn = MT8821C_Identify(addr)
    LogOp addr, "*IDN?", idn, log

    ' リセット
    Dim rst As String: rst = MT8821C_Reset(addr)
    LogOp addr, "*RST", rst, log

    ' プリセット
    Dim pre As String: pre = MT8821C_Preset(addr)
    LogOp addr, "PRESET", pre, log

    ' エラー確認
    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [GPIB] LTE 設定サンプル
' 実行内容: Band設定 -> Channel設定 -> DL Power設定 -> 読み返し
'=============================================================================
Public Sub Sample_GPIB_LteSetup()
    Const TITLE As String = "[GPIB] LTE 設定"
    Dim addr As String: addr = ADDR_GPIB
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' Band 設定 & 読み返し
    Dim setRes As String
    setRes = MT8821C_SetBand(addr, LTE_BAND)
    LogOp addr, "BAND " & LTE_BAND, setRes, log
    Dim band As String: band = MT8821C_GetBand(addr)
    LogOp addr, "BAND?", band, log

    ' Channel 設定 & 読み返し
    setRes = MT8821C_SetChannel(addr, LTE_CHANNEL)
    LogOp addr, "CHANL " & LTE_CHANNEL, setRes, log
    Dim ch As String: ch = MT8821C_GetChannel(addr)
    LogOp addr, "CHANL?", ch, log

    ' DL Power 設定 & 読み返し
    setRes = MT8821C_SetDlPower(addr, LTE_DL_POWER)
    LogOp addr, "OLVL " & LTE_DL_POWER, setRes, log
    Dim pwr As String: pwr = MT8821C_GetDlPower(addr)
    LogOp addr, "OLVL?", pwr, log

    ' エラー確認
    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'=============================================================================
' [GPIB] コール接続テスト
' 実行内容: コール接続 -> 接続状態確認 -> UL電力測定 -> 切断
'=============================================================================
Public Sub Sample_GPIB_CallTest()
    Const TITLE As String = "[GPIB] コール接続テスト"
    Dim addr As String: addr = ADDR_GPIB
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    If MsgBox("UE (端末) を接続してからOKを押してください。" & vbCrLf & _
              "接続先: " & addr, vbOKCancel + vbQuestion, TITLE) = vbCancel Then
        Exit Sub
    End If

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    ' コール接続
    Dim conn As String: conn = MT8821C_CallConnect(addr)
    LogOp addr, "CALLSO", conn, log

    ' 接続状態確認 (3秒待機)
    Application.Wait Now + TimeSerial(0, 0, 3)
    Dim stat As String: stat = MT8821C_GetCallStatus(addr)
    LogOp addr, "CALLSTAT?", stat, log

    ' UL 電力測定
    Application.StatusBar = "UL電力測定中..."
    Dim ulpwr As String: ulpwr = MT8821C_MeasureUlPower(addr)
    LogOp addr, "MEAS:UL:POW?", ulpwr, log

    ' コール切断
    Dim disc As String: disc = MT8821C_CallDisconnect(addr)
    LogOp addr, "CALLEND", disc, log

    ' エラー確認
    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'-----------------------------------------------------------------------------
' [GPIB] 全サンプルを順番に実行
'-----------------------------------------------------------------------------
Public Sub RunAllSamples_GPIB()
    Call Sample_GPIB_BasicCheck
    Call Sample_GPIB_LteSetup
    Call Sample_GPIB_CallTest
End Sub

'=============================================================================
' [LAN - VXI-11] 基本動作確認
'=============================================================================
Public Sub Sample_LAN_BasicCheck()
    Const TITLE As String = "[LAN VXI-11] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_VXI
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String: idn = MT8821C_Identify(addr)
    LogOp addr, "*IDN?", idn, log

    Dim rst As String: rst = MT8821C_Reset(addr)
    LogOp addr, "*RST", rst, log

    Dim pre As String: pre = MT8821C_Preset(addr)
    LogOp addr, "PRESET", pre, log

    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

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
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim setRes As String
    setRes = MT8821C_SetBand(addr, LTE_BAND)
    LogOp addr, "BAND " & LTE_BAND, setRes, log
    Dim band As String: band = MT8821C_GetBand(addr)
    LogOp addr, "BAND?", band, log

    setRes = MT8821C_SetChannel(addr, LTE_CHANNEL)
    LogOp addr, "CHANL " & LTE_CHANNEL, setRes, log
    Dim ch As String: ch = MT8821C_GetChannel(addr)
    LogOp addr, "CHANL?", ch, log

    setRes = MT8821C_SetDlPower(addr, LTE_DL_POWER)
    LogOp addr, "OLVL " & LTE_DL_POWER, setRes, log
    Dim pwr As String: pwr = MT8821C_GetDlPower(addr)
    LogOp addr, "OLVL?", pwr, log

    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

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
    Dim log  As String: log  = "接続先: " & addr & vbCrLf & vbCrLf

    If MsgBox("UE (端末) を接続してからOKを押してください。" & vbCrLf & _
              "接続先: " & addr, vbOKCancel + vbQuestion, TITLE) = vbCancel Then
        Exit Sub
    End If

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim conn As String: conn = MT8821C_CallConnect(addr)
    LogOp addr, "CALLSO", conn, log

    Application.Wait Now + TimeSerial(0, 0, 3)
    Dim stat As String: stat = MT8821C_GetCallStatus(addr)
    LogOp addr, "CALLSTAT?", stat, log

    Dim ulpwr As String: ulpwr = MT8821C_MeasureUlPower(addr)
    LogOp addr, "MEAS:UL:POW?", ulpwr, log

    Dim disc As String: disc = MT8821C_CallDisconnect(addr)
    LogOp addr, "CALLEND", disc, log

    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description, vbCritical, TITLE
End Sub

'-----------------------------------------------------------------------------
' [LAN - VXI-11] 全サンプルを順番に実行
'-----------------------------------------------------------------------------
Public Sub RunAllSamples_LAN()
    Call Sample_LAN_BasicCheck
    Call Sample_LAN_LteSetup
    Call Sample_LAN_CallTest
End Sub

'=============================================================================
' [LAN - Raw Socket] 基本動作確認
'=============================================================================
Public Sub Sample_Socket_BasicCheck()
    Const TITLE As String = "[LAN Socket] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_SOCK
    Dim log  As String
    log = "接続先: " & addr & vbCrLf & "(Raw Socket - ポート5025, 終端文字 LF)" & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String: idn = MT8821C_Identify(addr)
    LogOp addr, "*IDN?", idn, log

    Dim rst As String: rst = MT8821C_Reset(addr)
    LogOp addr, "*RST", rst, log

    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "※ MT8821C が Raw Socket 未対応の場合は VXI-11 を使用してください。", _
           vbCritical, TITLE
End Sub

'=============================================================================
' [LAN - HiSLIP] 基本動作確認
'=============================================================================
Public Sub Sample_HiSLIP_BasicCheck()
    Const TITLE As String = "[LAN HiSLIP] 基本動作確認"
    Dim addr As String: addr = ADDR_LAN_HISL
    Dim log  As String
    log = "接続先: " & addr & vbCrLf & "(HiSLIP - hislip0)" & vbCrLf & vbCrLf

    Application.StatusBar = TITLE & " 実行中..."
    On Error GoTo ErrHandler

    Dim idn As String: idn = MT8821C_Identify(addr)
    LogOp addr, "*IDN?", idn, log

    Dim rst As String: rst = MT8821C_Reset(addr)
    LogOp addr, "*RST", rst, log

    Dim err As String: err = MT8821C_GetError(addr)
    LogOp addr, "ERROR?", err, log

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "※ MT8821C が HiSLIP 未対応の場合は VXI-11 を使用してください。", _
           vbCritical, TITLE
End Sub

'=============================================================================
' 接続方式の比較確認
' GPIB / LAN (VXI-11 / Socket / HiSLIP) すべてで *IDN? を実行して比較する
'=============================================================================
Public Sub Sample_ConnectionComparison()
    Const TITLE As String = "接続方式 比較確認"
    Dim log As String
    log = "各接続方式での *IDN? 結果" & vbCrLf & String(44, "-") & vbCrLf & vbCrLf

    Dim addrs(3) As String
    Dim labels(3) As String
    addrs(0) = ADDR_GPIB:     labels(0) = "[GPIB]      "
    addrs(1) = ADDR_LAN_VXI:  labels(1) = "[LAN VXI-11]"
    addrs(2) = ADDR_LAN_SOCK: labels(2) = "[LAN Socket]"
    addrs(3) = ADDR_LAN_HISL: labels(3) = "[LAN HiSLIP]"

    Dim i As Integer
    For i = 0 To 3
        Application.StatusBar = labels(i) & " 確認中..."
        On Error Resume Next
        Dim response As String
        response = MT8821C_Identify(addrs(i))
        Dim isOK As Boolean
        isOK = (Err.Number = 0) And (InStr(UCase(response), "ERROR") = 0) And (response <> "")
        Err.Clear
        On Error GoTo 0

        Call Result_AppendRow(DEVICE_NAME, addrs(i), "*IDN?", response, isOK, "比較確認")
        log = log & labels(i) & vbCrLf
        log = log & "  addr: " & addrs(i) & vbCrLf
        log = log & "  IDN : " & IIf(isOK, response, "(失敗) " & response) & vbCrLf & vbCrLf
    Next i

    Application.StatusBar = False
    MsgBox log, vbInformation, TITLE
End Sub

'=============================================================================
' プライベート ヘルパー
'=============================================================================

'-----------------------------------------------------------------------------
' 1操作分の結果を Result シートへ記録し、ログ文字列にも追記する
'
' 引数:
'   addr    : VISAアドレス
'   command : コマンド名 (例: *IDN? / BAND 1 / CALLSO)
'   response: MT8821C_* 関数の戻り値
'   log     : MsgBox 用ログ文字列 (参照渡しで追記)
'-----------------------------------------------------------------------------
Private Sub LogOp(addr As String, command As String, response As String, ByRef log As String)
    Dim isQuery As Boolean: isQuery = (InStr(command, "?") > 0)
    Dim isOK    As Boolean

    If InStr(UCase(response), "ERROR") > 0 Then
        ' レスポンスに ERROR 文字列 -> 失敗
        isOK = False
    ElseIf isQuery Then
        ' クエリ: 応答が空でなければ成功
        isOK = (Len(Trim(response)) > 0)
    Else
        ' 書き込みコマンド: エラーがなければ成功 (応答は通常空)
        isOK = True
    End If

    Call Result_AppendRow(DEVICE_NAME, addr, command, response, isOK)

    ' MsgBox 用ログ
    If Not isQuery And response = "" Then
        log = log & command & ": OK" & vbCrLf
    Else
        log = log & command & ": " & response & vbCrLf
    End If
End Sub
