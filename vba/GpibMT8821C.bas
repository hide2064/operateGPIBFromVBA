Attribute VB_Name = "GpibMT8821C"
'=============================================================================
' GpibMT8821C.bas - Anritsu MT8821C 専用操作モジュール
'
' 【追加方法】 このファイルをVBAエディタからインポートするだけで機能が追加される
' 【削除方法】 このファイルをVBAプロジェクトから削除するだけで機能が無効化される
'
' 【前提】
'   - AppConfig.bas がインポートされていること
'   - Flaskサーバーが起動済みであること (start_server.bat を実行)
'
' 【使い方】
'   1. このファイルをVBAエディタからインポートする
'   2. MT8821C の GPIBアドレスを引数に渡して各関数を呼び出す
'   3. Control シートのボタンに割り当てる場合は
'      ExecuteMT8821CFromSheet() を使用する
'
' 【注意】
'   SCPI コマンドは MT8821C Operation Manual で要確認。
'   実機との接続前にコマンド構文を検証すること。
'=============================================================================
Option Explicit

' ===== MT8821C エンドポイント =====
Private Const MT8821C_PATH As String = "/mt8821c/execute"

'=============================================================================
' パブリック関数 - ボタンや他のモジュールから呼び出す
'=============================================================================

'-----------------------------------------------------------------------------
' 機器識別情報を取得する (*IDN?)
'-----------------------------------------------------------------------------
Public Function MT8821C_Identify(address As String) As String
    MT8821C_Identify = MT8821C_Call(address, "identify", "")
End Function

'-----------------------------------------------------------------------------
' システムリセット (*RST)
'-----------------------------------------------------------------------------
Public Function MT8821C_Reset(address As String) As String
    MT8821C_Reset = MT8821C_Call(address, "reset", "")
End Function

'-----------------------------------------------------------------------------
' システムプリセット (SYSTem:PRESet)
'-----------------------------------------------------------------------------
Public Function MT8821C_Preset(address As String) As String
    MT8821C_Preset = MT8821C_Call(address, "preset", "")
End Function

'-----------------------------------------------------------------------------
' エラー情報を取得する (SYSTem:ERRor?)
'-----------------------------------------------------------------------------
Public Function MT8821C_GetError(address As String) As String
    MT8821C_GetError = MT8821C_Call(address, "get_error", "")
End Function

'-----------------------------------------------------------------------------
' ダウンリンク出力レベルを取得する (BS:OLVL?)
'-----------------------------------------------------------------------------
Public Function MT8821C_GetDlPower(address As String) As String
    MT8821C_GetDlPower = MT8821C_Call(address, "get_dl_power", "")
End Function

'-----------------------------------------------------------------------------
' ダウンリンク出力レベルを設定する (BS:OLVL <power>)
' power_dbm: 出力レベル (dBm 単位、例: -70.0)
'-----------------------------------------------------------------------------
Public Function MT8821C_SetDlPower(address As String, power_dbm As Double) As String
    Dim params As String
    params = "{""power"": " & Format(power_dbm, "0.0") & "}"
    MT8821C_SetDlPower = MT8821C_Call(address, "set_dl_power", params)
End Function

'-----------------------------------------------------------------------------
' バンドを取得する (BAND?)
'-----------------------------------------------------------------------------
Public Function MT8821C_GetBand(address As String) As String
    MT8821C_GetBand = MT8821C_Call(address, "get_band", "")
End Function

'-----------------------------------------------------------------------------
' バンドを設定する (BAND <band>)
' band: バンド番号 (例: 1, 3, 7, 28)
'-----------------------------------------------------------------------------
Public Function MT8821C_SetBand(address As String, band As Integer) As String
    Dim params As String
    params = "{""band"": " & CStr(band) & "}"
    MT8821C_SetBand = MT8821C_Call(address, "set_band", params)
End Function

'-----------------------------------------------------------------------------
' チャネル番号を取得する (CHANL?)
'-----------------------------------------------------------------------------
Public Function MT8821C_GetChannel(address As String) As String
    MT8821C_GetChannel = MT8821C_Call(address, "get_channel", "")
End Function

'-----------------------------------------------------------------------------
' チャネル番号を設定する (CHANL <channel>)
' channel: チャネル番号 (例: 300)
'-----------------------------------------------------------------------------
Public Function MT8821C_SetChannel(address As String, channel As Long) As String
    Dim params As String
    params = "{""channel"": " & CStr(channel) & "}"
    MT8821C_SetChannel = MT8821C_Call(address, "set_channel", params)
End Function

'-----------------------------------------------------------------------------
' コール接続を開始する (CALLSO)
'-----------------------------------------------------------------------------
Public Function MT8821C_CallConnect(address As String) As String
    MT8821C_CallConnect = MT8821C_Call(address, "call_connect", "")
End Function

'-----------------------------------------------------------------------------
' コールを切断する (CALLEND)
'-----------------------------------------------------------------------------
Public Function MT8821C_CallDisconnect(address As String) As String
    MT8821C_CallDisconnect = MT8821C_Call(address, "call_disconnect", "")
End Function

'-----------------------------------------------------------------------------
' コール状態を取得する (CALLSTAT?)
'-----------------------------------------------------------------------------
Public Function MT8821C_GetCallStatus(address As String) As String
    MT8821C_GetCallStatus = MT8821C_Call(address, "get_call_status", "")
End Function

'-----------------------------------------------------------------------------
' アップリンク電力を測定する (MEAS:UL:POW?)
'-----------------------------------------------------------------------------
Public Function MT8821C_MeasureUlPower(address As String) As String
    MT8821C_MeasureUlPower = MT8821C_Call(address, "measure_ul_power", "")
End Function

'=============================================================================
' Controlシート連携 (ボタンに直接割り当て可能)
'=============================================================================

'-----------------------------------------------------------------------------
' Controlシートの選択行を MT8821C として実行する
' Controlシートの B列 に "アクション名" を入力しておく
' (例: identify / preset / call_connect / set_dl_power etc.)
'
' パラメータが必要なアクション (set_dl_power 等) は
' E列に JSON 形式で入力する (例: {"power": -70.0})
'-----------------------------------------------------------------------------
Public Sub ExecuteMT8821CFromSheet()
    Dim wsControl As Worksheet
    Dim wsConfig As Worksheet
    Dim selectedRow As Long
    Dim deviceName As String
    Dim action As String
    Dim paramsJson As String
    Dim address As String
    Dim timeout As Long
    Dim jsonResult As String
    Dim success As Boolean
    Dim response As String
    Dim errorMsg As String

    On Error GoTo ErrorHandler

    wsControl = ThisWorkbook.Worksheets("Control")
    wsConfig  = ThisWorkbook.Worksheets("Config")
    selectedRow = ActiveCell.Row

    If selectedRow < 2 Then
        MsgBox "2行目以降のコマンド行を選択してください。", vbExclamation
        Exit Sub
    End If

    deviceName = Trim(wsControl.Cells(selectedRow, 1).Value)  ' A列: 機器名
    action     = Trim(wsControl.Cells(selectedRow, 2).Value)  ' B列: アクション名
    paramsJson = Trim(wsControl.Cells(selectedRow, 5).Value)  ' E列: パラメータJSON (省略可)

    If deviceName = "" Or action = "" Then
        MsgBox "機器名とアクション名を入力してください。", vbExclamation
        Exit Sub
    End If

    ' Config シートからアドレスとタイムアウトを取得
    If Not GetDeviceAddress(wsConfig, deviceName, address, timeout) Then
        wsControl.Cells(selectedRow, 4).Value = "ERROR: 機器 '" & deviceName & "' がConfigシートに見つかりません"
        wsControl.Cells(selectedRow, 4).Font.Color = RGB(255, 0, 0)
        Exit Sub
    End If

    Application.StatusBar = "MT8821C 実行中: " & deviceName & " / " & action
    jsonResult = MT8821C_Call(address, action, paramsJson)
    ParseJsonResult jsonResult, success, response, errorMsg

    wsControl.Cells(selectedRow, 3).Value = response           ' C列: 応答結果
    If success Then
        wsControl.Cells(selectedRow, 4).Value = "OK"
        wsControl.Cells(selectedRow, 4).Font.Color = RGB(0, 128, 0)
    Else
        wsControl.Cells(selectedRow, 4).Value = "ERROR: " & errorMsg
        wsControl.Cells(selectedRow, 4).Font.Color = RGB(255, 0, 0)
    End If

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
' MT8821C Blueprint の /execute エンドポイントを呼び出す
' 戻り値: レスポンスのJSON文字列
'-----------------------------------------------------------------------------
Private Function MT8821C_Call(address As String, action As String, paramsJson As String) As String
    Dim http As Object
    Dim url As String
    Dim body As String

    url = AppConfig.ServerBaseUrl() & MT8821C_PATH

    If paramsJson = "" Then
        body = "{""address"": """ & JsonEscape(address) & """" _
             & ", ""action"": """ & action & """}"
    Else
        body = "{""address"": """ & JsonEscape(address) & """" _
             & ", ""action"": """ & action & """" _
             & ", ""params"": " & paramsJson & "}"
    End If

    On Error GoTo HttpError
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send body
    MT8821C_Call = http.ResponseText
    Exit Function

HttpError:
    MT8821C_Call = "{""success"": false, ""response"": """", ""error"": ""HTTPエラー: " & JsonEscape(Err.Description) & """}"
End Function

'-----------------------------------------------------------------------------
' Config シートから機器アドレスとタイムアウトを取得する
'-----------------------------------------------------------------------------
Private Function GetDeviceAddress(wsConfig As Worksheet, deviceName As String, ByRef address As String, ByRef timeout As Long) As Boolean
    Dim lastRow As Long
    Dim i As Long

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(wsConfig.Cells(i, 1).Value) = deviceName Then
            address = Trim(wsConfig.Cells(i, 2).Value)
            timeout = CLng(wsConfig.Cells(i, 3).Value)
            If timeout <= 0 Then timeout = 5000
            GetDeviceAddress = True
            Exit Function
        End If
    Next i
    GetDeviceAddress = False
End Function

'-----------------------------------------------------------------------------
' JSON 文字列から各フィールドを解析する
'-----------------------------------------------------------------------------
Private Sub ParseJsonResult(jsonStr As String, ByRef success As Boolean, ByRef response As String, ByRef errorMsg As String)
    success  = (InStr(jsonStr, """success"": true") > 0)
    response = ExtractJsonString(jsonStr, "response")
    errorMsg = ExtractJsonString(jsonStr, "error")
    If Not success And errorMsg = "" Then errorMsg = "応答を解析できませんでした"
End Sub

Private Function ExtractJsonString(jsonStr As String, key As String) As String
    Dim pattern As String
    Dim startPos As Long
    Dim endPos As Long
    pattern  = """" & key & """: """
    startPos = InStr(jsonStr, pattern)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(pattern)
    endPos   = InStr(startPos, jsonStr, """")
    If endPos = 0 Then Exit Function
    ExtractJsonString = Mid(jsonStr, startPos, endPos - startPos)
End Function

Private Function JsonEscape(str As String) As String
    Dim s As String
    s = Replace(str, "\", "\\")
    s = Replace(s, """", "\""")
    JsonEscape = s
End Function
