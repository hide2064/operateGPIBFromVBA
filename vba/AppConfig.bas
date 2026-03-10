Attribute VB_Name = "AppConfig"
'=============================================================================
' AppConfig.bas - アプリケーション設定読み込みモジュール
'
' config\settings.ini を Windows API 経由で読み込む。
' GpibControlHttp.bas など他のモジュールから呼び出して使う。
'
' INIファイルのパス: ThisWorkbook.Path\config\settings.ini
'   (Excelファイルと同じドライブ・フォルダ構成に配置すること)
'=============================================================================
Option Explicit

' Windows API: INIファイルの文字列値を読み込む
#If Win64 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#End If

Private Const BUFFER_SIZE As Long = 512
Private Const CONFIG_REL_PATH As String = "\config\settings.ini"

'=============================================================================
' 基本読み込み関数
'=============================================================================

' INIファイルの絶対パスを返す
Public Function ConfigFilePath() As String
    ConfigFilePath = ThisWorkbook.Path & CONFIG_REL_PATH
End Function

' 文字列値を取得する
Public Function GetString(section As String, key As String, defaultValue As String) As String
    Dim buf As String
    Dim ret As Long
    buf = String(BUFFER_SIZE, Chr(0))
    ret = GetPrivateProfileString(section, key, defaultValue, buf, BUFFER_SIZE, ConfigFilePath())
    GetString = Left(buf, ret)
End Function

' 整数値を取得する
Public Function GetInt(section As String, key As String, defaultValue As Long) As Long
    Dim val As String
    val = GetString(section, key, CStr(defaultValue))
    If IsNumeric(val) Then
        GetInt = CLng(val)
    Else
        GetInt = defaultValue
    End If
End Function

'=============================================================================
' [Server] セクション
'=============================================================================

' FlaskサーバーのベースURL (例: http://127.0.0.1:5000)
Public Function ServerBaseUrl() As String
    Dim host As String
    Dim port As Long
    host = GetString("Server", "Host", "127.0.0.1")
    port = GetInt("Server", "Port", 5000)
    ServerBaseUrl = "http://" & host & ":" & CStr(port)
End Function

' Python実行ファイルのパス
Public Function PythonExe() As String
    PythonExe = GetString("Server", "PythonExe", "python")
End Function

' server.py の絶対パス
Public Function ServerScript() As String
    ServerScript = GetString("Server", "ServerScript", "")
End Function

' サーバー起動確認のタイムアウト秒数
Public Function HealthTimeoutSec() As Long
    HealthTimeoutSec = GetInt("Server", "HealthTimeoutSec", 10)
End Function

'=============================================================================
' [Lan] セクション
'=============================================================================

' Raw Socket 接続のデフォルトポート
Public Function DefaultSocketPort() As Long
    DefaultSocketPort = GetInt("Lan", "DefaultSocketPort", 5025)
End Function

'=============================================================================
' VISAアドレスビルダー
'
' Config シートの D列(Protocol) / E列(Host) / F列(Port) から
' VISA アドレス文字列を組み立てる。
' B列にフルアドレスが入力されている場合はそちらを優先すること。
'
' 対応プロトコル:
'   GPIB   → GPIB0::<host>::INSTR          (<host> は GPIB アドレス番号)
'   TCPIP  → TCPIP0::<host>::INSTR         (VXI-11 標準)
'   SOCKET → TCPIP0::<host>::<port>::SOCKET (Raw TCP)
'   HISLIP → TCPIP0::<host>::hislip0::INSTR (高速 LAN)
'=============================================================================

Public Function BuildVisaAddress(protocol As String, host As String, Optional port As String = "") As String
    Dim p As String
    p = UCase(Trim(protocol))

    Select Case p
        Case "GPIB"
            BuildVisaAddress = "GPIB0::" & Trim(host) & "::INSTR"

        Case "SOCKET", "TCPIP_SOCKET"
            Dim actualPort As String
            actualPort = Trim(port)
            If actualPort = "" Then actualPort = CStr(DefaultSocketPort())
            BuildVisaAddress = "TCPIP0::" & Trim(host) & "::" & actualPort & "::SOCKET"

        Case "HISLIP", "TCPIP_HISLIP"
            BuildVisaAddress = "TCPIP0::" & Trim(host) & "::hislip0::INSTR"

        Case "TCPIP", "VXI11", "LAN", "TCPIP_VXI11"
            BuildVisaAddress = "TCPIP0::" & Trim(host) & "::INSTR"

        Case Else
            ' 不明なプロトコル: host をそのまま VISA アドレスとして扱う
            BuildVisaAddress = Trim(host)
    End Select
End Function

'=============================================================================
' デバッグ用: 設定内容をメッセージボックスに表示する
'=============================================================================
Public Sub ShowConfig()
    MsgBox "【現在の設定】" & vbCrLf & vbCrLf & _
           "INIファイル: " & ConfigFilePath() & vbCrLf & vbCrLf & _
           "[Server]" & vbCrLf & _
           "  ServerBaseUrl : " & ServerBaseUrl() & vbCrLf & _
           "  PythonExe     : " & PythonExe() & vbCrLf & _
           "  ServerScript  : " & ServerScript() & vbCrLf & _
           "  HealthTimeout : " & HealthTimeoutSec() & " 秒" & vbCrLf & vbCrLf & _
           "[Lan]" & vbCrLf & _
           "  DefaultSocketPort : " & DefaultSocketPort(), _
           vbInformation, "AppConfig"
End Sub
