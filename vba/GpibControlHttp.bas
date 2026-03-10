Attribute VB_Name = "GpibControlHttp"
'=============================================================================
' GpibControlHttp.bas - Flask経由GPIB制御モジュール
'
' GpibControl.bas (CLI方式) の追加モジュール。
' Pythonをプロセス都度起動する代わりに、常駐するFlaskサーバーにHTTPで
' リクエストを投げる方式。接続プール・リトライはサーバー側で処理される。
'
' 【GpibControl.bas との使い分け】
'   GpibControl.bas    : Pythonを都度起動。サーバー不要だが低速・重い。
'   GpibControlHttp.bas: Flaskサーバー常駐。高速・安定。本番運用向け。
'
' 【前提】
'   Flaskサーバーが起動済みであること。
'   起動方法: start_server.bat を実行、または以下をコマンドプロンプトで実行
'     python C:\work\operateGPIBFromVBA\operateGPIBFromVBA\python\server.py
'
' 【使い方】
'   1. このファイルをVBAエディタからインポートする
'   2. SERVER_BASE_URL を環境に合わせて設定する (デフォルトは localhost:5000)
'   3. Controlシートにボタンを作成し、以下のマクロを割り当てる:
'      - ExecuteAllCommandsHttp    (すべて実行)
'      - ExecuteSelectedCommandHttp (選択行を実行)
'      - StartGpibServer           (サーバー起動)
'=============================================================================
Option Explicit

' ===== 設定定数 =====
Private Const SERVER_BASE_URL As String = "http://127.0.0.1:5000"
Private Const PYTHON_EXE As String = "python"
Private Const SERVER_SCRIPT As String = "C:\work\operateGPIBFromVBA\operateGPIBFromVBA\python\server.py"
Private Const HEALTH_TIMEOUT_SEC As Integer = 10  ' サーバー起動待ちの最大秒数

' ===== シート名定数 (GpibControl.bas と共通) =====
Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_CONTROL As String = "Control"

' ===== Configシートの列番号 =====
Private Const COL_CFG_NAME As Integer = 1
Private Const COL_CFG_ADDRESS As Integer = 2
Private Const COL_CFG_TIMEOUT As Integer = 3

' ===== Controlシートの列番号 =====
Private Const COL_CTL_NAME As Integer = 1
Private Const COL_CTL_COMMAND As Integer = 2
Private Const COL_CTL_RESPONSE As Integer = 3
Private Const COL_CTL_STATUS As Integer = 4

'=============================================================================
' パブリック関数 (ボタンに割り当てる)
'=============================================================================

'-----------------------------------------------------------------------------
' Controlシートのすべてのコマンドを実行する
'-----------------------------------------------------------------------------
Public Sub ExecuteAllCommandsHttp()
    Dim wsControl As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    If Not EnsureServerRunning() Then Exit Sub

    wsControl = GetSheet(SHEET_CONTROL)
    lastRow = GetLastRow(wsControl, COL_CTL_NAME)

    If lastRow < 2 Then
        MsgBox "実行するコマンドがありません。", vbInformation
        Exit Sub
    End If

    Application.StatusBar = "GPIB コマンドを実行中 (Flask経由)..."

    For i = 2 To lastRow
        Dim deviceName As String
        Dim command As String
        deviceName = Trim(wsControl.Cells(i, COL_CTL_NAME).Value)
        command = Trim(wsControl.Cells(i, COL_CTL_COMMAND).Value)
        If deviceName = "" Or command = "" Then GoTo NextRow
        Call ExecuteSingleCommandHttp(wsControl, i, deviceName, command)
NextRow:
    Next i

    Application.StatusBar = "GPIB コマンドの実行が完了しました (Flask経由)。"
    MsgBox "すべてのコマンドを実行しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

'-----------------------------------------------------------------------------
' Controlシートで選択中の行のコマンドを実行する
'-----------------------------------------------------------------------------
Public Sub ExecuteSelectedCommandHttp()
    Dim wsControl As Worksheet
    Dim selectedRow As Long
    Dim deviceName As String
    Dim command As String

    On Error GoTo ErrorHandler

    If Not EnsureServerRunning() Then Exit Sub

    wsControl = GetSheet(SHEET_CONTROL)
    selectedRow = ActiveCell.Row

    If selectedRow < 2 Then
        MsgBox "2行目以降のコマンド行を選択してください。", vbExclamation
        Exit Sub
    End If

    deviceName = Trim(wsControl.Cells(selectedRow, COL_CTL_NAME).Value)
    command = Trim(wsControl.Cells(selectedRow, COL_CTL_COMMAND).Value)

    If deviceName = "" Or command = "" Then
        MsgBox "機器名とコマンドを入力してください。", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "GPIB コマンドを実行中 (Flask経由): " & deviceName & " / " & command
    Call ExecuteSingleCommandHttp(wsControl, selectedRow, deviceName, command)
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

'-----------------------------------------------------------------------------
' Flaskサーバーをバックグラウンドで起動する
'-----------------------------------------------------------------------------
Public Sub StartGpibServer()
    If IsServerRunning() Then
        MsgBox "サーバーはすでに起動しています。" & vbCrLf & SERVER_BASE_URL, vbInformation
        Exit Sub
    End If

    Dim cmd As String
    cmd = "cmd /c start /min """ & PYTHON_EXE & """ """ & SERVER_SCRIPT & """"
    Shell cmd, vbHide

    Application.StatusBar = "サーバーを起動中..."
    If WaitForServer(HEALTH_TIMEOUT_SEC) Then
        Application.StatusBar = False
        MsgBox "サーバーが起動しました。" & vbCrLf & SERVER_BASE_URL, vbInformation
    Else
        Application.StatusBar = False
        MsgBox "サーバーの起動を確認できませんでした。" & vbCrLf & _
               "手動で起動してください: python server.py", vbExclamation
    End If
End Sub

'=============================================================================
' プライベート関数
'=============================================================================

'-----------------------------------------------------------------------------
' 1件のGPIBコマンドを実行してシートに結果を書き込む
'-----------------------------------------------------------------------------
Private Sub ExecuteSingleCommandHttp(ws As Worksheet, rowIndex As Long, deviceName As String, command As String)
    Dim address As String
    Dim timeout As Long
    Dim jsonResult As String
    Dim success As Boolean
    Dim response As String
    Dim errorMsg As String

    If Not GetDeviceConfig(deviceName, address, timeout) Then
        ws.Cells(rowIndex, COL_CTL_STATUS).Value = "ERROR: 機器 '" & deviceName & "' がConfigシートに見つかりません"
        ws.Cells(rowIndex, COL_CTL_STATUS).Font.Color = RGB(255, 0, 0)
        Exit Sub
    End If

    jsonResult = PostExecute(address, command, timeout)
    ParseJsonResult jsonResult, success, response, errorMsg

    ws.Cells(rowIndex, COL_CTL_RESPONSE).Value = response
    If success Then
        ws.Cells(rowIndex, COL_CTL_STATUS).Value = "OK"
        ws.Cells(rowIndex, COL_CTL_STATUS).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(rowIndex, COL_CTL_STATUS).Value = "ERROR: " & errorMsg
        ws.Cells(rowIndex, COL_CTL_STATUS).Font.Color = RGB(255, 0, 0)
    End If
End Sub

'-----------------------------------------------------------------------------
' POST /execute を呼び出す
' 戻り値: レスポンスのJSON文字列
'-----------------------------------------------------------------------------
Private Function PostExecute(address As String, command As String, timeout As Long) As String
    Dim http As Object
    Dim url As String
    Dim body As String

    url = SERVER_BASE_URL & "/execute"
    body = "{""address"": """ & JsonEscape(address) & """" _
         & ", ""command"": """ & JsonEscape(command) & """" _
         & ", ""timeout"": " & CStr(timeout) & "}"

    On Error GoTo HttpError
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send body
    PostExecute = http.ResponseText
    Exit Function

HttpError:
    PostExecute = "{""success"": false, ""response"": """", ""error"": ""HTTPリクエスト失敗: " & JsonEscape(Err.Description) & """}"
End Function

'-----------------------------------------------------------------------------
' GET /health でサーバーの稼働を確認する
'-----------------------------------------------------------------------------
Private Function IsServerRunning() As Boolean
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", SERVER_BASE_URL & "/health", False
    http.setTimeouts 0, 1000, 1000, 1000
    http.Send
    IsServerRunning = (Err.Number = 0 And http.Status = 200)
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------------
' サーバーが起動するまで待つ (最大 maxSec 秒)
'-----------------------------------------------------------------------------
Private Function WaitForServer(maxSec As Integer) As Boolean
    Dim startTime As Single
    startTime = Timer
    Do
        If IsServerRunning() Then
            WaitForServer = True
            Exit Function
        End If
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop While Timer - startTime < maxSec
    WaitForServer = False
End Function

'-----------------------------------------------------------------------------
' サーバーが起動していなければ確認ダイアログを出す
'-----------------------------------------------------------------------------
Private Function EnsureServerRunning() As Boolean
    If IsServerRunning() Then
        EnsureServerRunning = True
        Exit Function
    End If

    Dim ans As VbMsgBoxResult
    ans = MsgBox("Flaskサーバーが起動していません。" & vbCrLf & _
                 "今すぐ起動しますか？", vbYesNo + vbQuestion)
    If ans = vbYes Then
        StartGpibServer
        EnsureServerRunning = IsServerRunning()
    Else
        EnsureServerRunning = False
    End If
End Function

'-----------------------------------------------------------------------------
' Configシートから機器のVISAアドレスとタイムアウトを取得する
'-----------------------------------------------------------------------------
Private Function GetDeviceConfig(deviceName As String, ByRef address As String, ByRef timeout As Long) As Boolean
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim i As Long

    wsConfig = GetSheet(SHEET_CONFIG)
    lastRow = GetLastRow(wsConfig, COL_CFG_NAME)

    For i = 2 To lastRow
        If Trim(wsConfig.Cells(i, COL_CFG_NAME).Value) = deviceName Then
            address = Trim(wsConfig.Cells(i, COL_CFG_ADDRESS).Value)
            timeout = CLng(wsConfig.Cells(i, COL_CFG_TIMEOUT).Value)
            If timeout <= 0 Then timeout = 5000
            GetDeviceConfig = True
            Exit Function
        End If
    Next i

    GetDeviceConfig = False
End Function

'-----------------------------------------------------------------------------
' JSON文字列から各フィールドを解析する
'-----------------------------------------------------------------------------
Private Sub ParseJsonResult(jsonStr As String, ByRef success As Boolean, ByRef response As String, ByRef errorMsg As String)
    success = False
    response = ""
    errorMsg = "サーバーからの応答を解析できませんでした"

    If jsonStr = "" Then
        errorMsg = "サーバーからの応答がありません"
        Exit Sub
    End If

    If InStr(jsonStr, """success"": true") > 0 Then
        success = True
    End If

    response = ExtractJsonString(jsonStr, "response")
    errorMsg = ExtractJsonString(jsonStr, "error")
End Sub

'-----------------------------------------------------------------------------
' JSON文字列から指定キーの値を取得する (簡易)
'-----------------------------------------------------------------------------
Private Function ExtractJsonString(jsonStr As String, key As String) As String
    Dim pattern As String
    Dim startPos As Long
    Dim endPos As Long

    pattern = """" & key & """: """
    startPos = InStr(jsonStr, pattern)
    If startPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If

    startPos = startPos + Len(pattern)
    endPos = InStr(startPos, jsonStr, """")
    If endPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If

    ExtractJsonString = Mid(jsonStr, startPos, endPos - startPos)
End Function

'-----------------------------------------------------------------------------
' JSON文字列用のエスケープ
'-----------------------------------------------------------------------------
Private Function JsonEscape(str As String) As String
    Dim s As String
    s = str
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEscape = s
End Function

'-----------------------------------------------------------------------------
' シート名からWorksheetオブジェクトを取得する
'-----------------------------------------------------------------------------
Private Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Err.Raise vbObjectError + 1, , "シート '" & sheetName & "' が見つかりません。"
    End If
End Function

'-----------------------------------------------------------------------------
' 指定列の最終行番号を取得する
'-----------------------------------------------------------------------------
Private Function GetLastRow(ws As Worksheet, col As Integer) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
