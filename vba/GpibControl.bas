Attribute VB_Name = "GpibControl"
'=============================================================================
' GpibControl.bas - GPIB制御モジュール
'
' 【使い方】
'   1. このファイルをExcelのVBAエディタからインポートする
'      (VBAエディタ > ファイル > ファイルのインポート)
'   2. 定数 PYTHON_EXE, SCRIPT_PATH を環境に合わせて設定する
'   3. Excelに以下のシートを作成する:
'      - "Config" シート : GPIB機器の設定情報
'      - "Control" シート: 操作ボタンと結果表示
'
' 【Configシートの構成】
'   行1: ヘッダー (Name / Address / Timeout)
'   行2以降: 機器ごとの設定
'   例:
'     A列(Name)    B列(Address)          C列(Timeout)
'     DMM          GPIB0::22::INSTR      5000
'     PowerSupply  GPIB0::5::INSTR       3000
'
' 【Controlシートの構成】
'   A列: 機器名 (ConfigシートのA列と対応)
'   B列: SCPIコマンド
'   C列: 実行結果 (自動入力)
'   D列: ステータス (自動入力)
'=============================================================================
Option Explicit

' ===== 設定定数 (環境に合わせて変更) =====
Private Const PYTHON_EXE As String = "python"        ' Pythonの実行ファイルパス
Private Const SCRIPT_PATH As String = "C:\work\operateGPIBFromVBA\operateGPIBFromVBA\python\gpib_controller.py"

' ===== シート名定数 =====
Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_CONTROL As String = "Control"

' ===== Configシートの列番号 =====
Private Const COL_CFG_NAME As Integer = 1     ' A列: 機器名
Private Const COL_CFG_ADDRESS As Integer = 2  ' B列: VISAアドレス
Private Const COL_CFG_TIMEOUT As Integer = 3  ' C列: タイムアウト(ms)

' ===== Controlシートの列番号 =====
Private Const COL_CTL_NAME As Integer = 1     ' A列: 機器名
Private Const COL_CTL_COMMAND As Integer = 2  ' B列: SCPIコマンド
Private Const COL_CTL_RESPONSE As Integer = 3 ' C列: 応答結果
Private Const COL_CTL_STATUS As Integer = 4   ' D列: ステータス

'=============================================================================
' パブリック関数
'=============================================================================

'-----------------------------------------------------------------------------
' ControlシートのアクティブなコマンドをすべてGPIBに送信する
' (Controlシートの「すべて実行」ボタンに割り当てる)
'-----------------------------------------------------------------------------
Public Sub ExecuteAllCommands()
    Dim wsControl As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    wsControl = GetSheet(SHEET_CONTROL)
    lastRow = GetLastRow(wsControl, COL_CTL_NAME)

    If lastRow < 2 Then
        MsgBox "実行するコマンドがありません。", vbInformation
        Exit Sub
    End If

    Application.StatusBar = "GPIB コマンドを実行中..."

    For i = 2 To lastRow
        Dim deviceName As String
        Dim command As String
        deviceName = Trim(wsControl.Cells(i, COL_CTL_NAME).Value)
        command = Trim(wsControl.Cells(i, COL_CTL_COMMAND).Value)

        If deviceName = "" Or command = "" Then
            GoTo NextRow
        End If

        Call ExecuteSingleCommand(wsControl, i, deviceName, command)
NextRow:
    Next i

    Application.StatusBar = "GPIB コマンドの実行が完了しました。"
    MsgBox "すべてのコマンドを実行しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

'-----------------------------------------------------------------------------
' Controlシートで選択中の行のコマンドを実行する
' (Controlシートの「選択行を実行」ボタンに割り当てる)
'-----------------------------------------------------------------------------
Public Sub ExecuteSelectedCommand()
    Dim wsControl As Worksheet
    Dim selectedRow As Long
    Dim deviceName As String
    Dim command As String

    On Error GoTo ErrorHandler

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

    Application.StatusBar = "GPIB コマンドを実行中: " & deviceName & " / " & command
    Call ExecuteSingleCommand(wsControl, selectedRow, deviceName, command)
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

'=============================================================================
' プライベート関数
'=============================================================================

'-----------------------------------------------------------------------------
' 指定行のGPIBコマンドを1件実行し、結果をシートに書き込む
'-----------------------------------------------------------------------------
Private Sub ExecuteSingleCommand(ws As Worksheet, rowIndex As Long, deviceName As String, command As String)
    Dim address As String
    Dim timeout As Long
    Dim jsonResult As String
    Dim success As Boolean
    Dim response As String
    Dim errorMsg As String

    ' ConfigシートからVISAアドレスとタイムアウトを取得
    If Not GetDeviceConfig(deviceName, address, timeout) Then
        ws.Cells(rowIndex, COL_CTL_RESPONSE).Value = ""
        ws.Cells(rowIndex, COL_CTL_STATUS).Value = "ERROR: 機器 '" & deviceName & "' がConfigシートに見つかりません"
        ws.Cells(rowIndex, COL_CTL_STATUS).Font.Color = RGB(255, 0, 0)
        Exit Sub
    End If

    ' Pythonスクリプトを呼び出す
    jsonResult = CallPythonGpib(address, command, timeout)

    ' JSONを解析して結果を取得
    ParseJsonResult jsonResult, success, response, errorMsg

    ' 結果をシートに書き込む
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
' Configシートから機器のVISAアドレスとタイムアウトを取得する
' 戻り値: 見つかった場合True、見つからない場合False
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
' Pythonスクリプトを呼び出してGPIBコマンドを実行する
' 戻り値: PythonのstdoutのJSON文字列
'-----------------------------------------------------------------------------
Private Function CallPythonGpib(address As String, command As String, timeout As Long) As String
    Dim shell As Object
    Dim exec As Object
    Dim cmd As String
    Dim output As String

    ' コマンドライン構築 (アドレスとコマンドをダブルクォートで囲む)
    cmd = PYTHON_EXE & " """ & SCRIPT_PATH & """" _
        & " --address """ & address & """" _
        & " --command """ & EscapeForCmd(command) & """" _
        & " --timeout " & CStr(timeout)

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(cmd)

    ' stdoutを読み取る
    output = exec.StdOut.ReadAll()

    ' プロセス終了を待つ
    Do While exec.Status = 0
        DoEvents
    Loop

    CallPythonGpib = Trim(output)
End Function

'-----------------------------------------------------------------------------
' JSON文字列から各フィールドを取得する (簡易パーサー)
' {"success": true/false, "response": "...", "error": "..."}
'-----------------------------------------------------------------------------
Private Sub ParseJsonResult(jsonStr As String, ByRef success As Boolean, ByRef response As String, ByRef errorMsg As String)
    success = False
    response = ""
    errorMsg = "Pythonからの応答を解析できませんでした"

    If jsonStr = "" Then
        errorMsg = "Pythonスクリプトからの応答がありません"
        Exit Sub
    End If

    ' "success": true / false の判定
    If InStr(jsonStr, """success"": true") > 0 Then
        success = True
    End If

    ' "response": "..." の取得
    response = ExtractJsonString(jsonStr, "response")

    ' "error": "..." の取得
    errorMsg = ExtractJsonString(jsonStr, "error")
End Sub

'-----------------------------------------------------------------------------
' JSON文字列から指定キーの文字列値を取得する (簡易)
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
' コマンドラインに渡す文字列のエスケープ (ダブルクォートをエスケープ)
'-----------------------------------------------------------------------------
Private Function EscapeForCmd(str As String) As String
    EscapeForCmd = Replace(str, """", "\""")
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
