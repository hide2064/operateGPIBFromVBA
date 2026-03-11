Attribute VB_Name = "ResultSheet"
'=============================================================================
' ResultSheet.bas - 試験結果 Result シート管理モジュール
'
' 【機能】
'   - MT8821C_Sample.bas など各サンプルから Result シートへ結果を1行ずつ追記する
'   - Control シートの実行済み結果を一括転記する
'   - Result シートの全データをクリアする
'
' 【前提】
'   - AppConfig.bas がインポートされていること
'   - Excel ファイルに「Result」シートがあること
'     (ない場合は自動作成する)
'
' 【Result シート列構成】
'   A: No. / B: 実行日時 / C: 機器名 / D: VISAアドレス / E: 接続方式
'   F: コマンド / G: 応答結果 / H: ステータス / I: 備考
'=============================================================================
Option Explicit

Private Const SHEET_RESULT As String = "Result"

' 列番号
Private Const COL_NO       As Integer = 1  ' A: No.
Private Const COL_DATETIME As Integer = 2  ' B: 実行日時
Private Const COL_DEVICE   As Integer = 3  ' C: 機器名
Private Const COL_ADDRESS  As Integer = 4  ' D: VISAアドレス
Private Const COL_CONNTYPE As Integer = 5  ' E: 接続方式
Private Const COL_COMMAND  As Integer = 6  ' F: コマンド / アクション
Private Const COL_RESPONSE As Integer = 7  ' G: 応答結果
Private Const COL_STATUS   As Integer = 8  ' H: ステータス
Private Const COL_NOTES    As Integer = 9  ' I: 備考

'=============================================================================
' パブリック関数 - ボタンや他のモジュールから呼び出す
'=============================================================================

'-----------------------------------------------------------------------------
' Result シートへ1行追記する
'
' 引数:
'   deviceName : 機器名 (例: MT8821C)
'   address    : VISAアドレス (例: GPIB0::1::INSTR)
'   command    : 実行コマンドまたはアクション名 (例: *IDN?)
'   response   : 機器からの応答結果
'   isSuccess  : 成功=True / 失敗=False
'   notes      : 備考 (省略可)
'-----------------------------------------------------------------------------
Public Sub Result_AppendRow(deviceName As String, _
                             address As String, _
                             command As String, _
                             response As String, _
                             isSuccess As Boolean, _
                             Optional notes As String = "")
    Dim ws As Worksheet
    Set ws = GetOrCreateResultSheet()

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ' データ書き込み
    ws.Cells(nextRow, COL_NO).Value       = nextRow - 1
    ws.Cells(nextRow, COL_DATETIME).Value = Now()
    ws.Cells(nextRow, COL_DEVICE).Value   = deviceName
    ws.Cells(nextRow, COL_ADDRESS).Value  = address
    ws.Cells(nextRow, COL_CONNTYPE).Value = GetConnectionType(address)
    ws.Cells(nextRow, COL_COMMAND).Value  = command
    ws.Cells(nextRow, COL_RESPONSE).Value = response
    ws.Cells(nextRow, COL_STATUS).Value   = IIf(isSuccess, "OK", "ERROR")
    ws.Cells(nextRow, COL_NOTES).Value    = notes

    ' 日時フォーマット
    ws.Cells(nextRow, COL_DATETIME).NumberFormat = "yyyy/mm/dd hh:mm:ss"

    ' ステータスセルの色
    With ws.Cells(nextRow, COL_STATUS).Font
        .Bold = True
        .Color = IIf(isSuccess, RGB(0, 128, 0), RGB(204, 0, 0))
    End With

    ' ゼブラストライプ (偶数行に薄い青)
    If nextRow Mod 2 = 0 Then
        Dim col As Integer
        For col = COL_NO To COL_NOTES
            ws.Cells(nextRow, col).Interior.Color = RGB(242, 247, 251)
        Next col
    End If
End Sub

'-----------------------------------------------------------------------------
' Control シートの実行済み結果を Result シートへ一括転記する
' ボタンに割り当てて使用する
'-----------------------------------------------------------------------------
Public Sub Result_AppendFromControl()
    Dim wsControl As Worksheet
    Dim wsConfig  As Worksheet

    On Error Resume Next
    Set wsControl = ThisWorkbook.Worksheets("Control")
    Set wsConfig  = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0

    If wsControl Is Nothing Then
        MsgBox "Control シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsControl.Cells(wsControl.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Control シートに実行済みデータがありません。", vbInformation
        Exit Sub
    End If

    Dim count As Long
    count = 0
    Dim i As Long

    For i = 2 To lastRow
        Dim deviceName As String: deviceName = Trim(wsControl.Cells(i, 1).Value)
        Dim command    As String: command    = Trim(wsControl.Cells(i, 2).Value)
        Dim response   As String: response   = Trim(wsControl.Cells(i, 3).Value)
        Dim status     As String: status     = Trim(wsControl.Cells(i, 4).Value)

        ' 未入力・未実行行はスキップ
        If deviceName = "" Or command = "" Or status = "" Then GoTo NextRow

        ' Config シートからアドレスを取得
        Dim address As String
        address = ""
        If Not wsConfig Is Nothing Then
            address = FindAddressByDevice(wsConfig, deviceName)
        End If

        Dim isSuccess As Boolean
        isSuccess = (status = "OK")

        Call Result_AppendRow(deviceName, address, command, response, isSuccess, "Control転記")
        count = count + 1
NextRow:
    Next i

    If count = 0 Then
        MsgBox "転記対象の実行済み行がありませんでした。" & vbCrLf & _
               "(ステータス列が空の行はスキップされます)", vbInformation
    Else
        MsgBox count & " 件の結果を Result シートに転記しました。", vbInformation
    End If
End Sub

'-----------------------------------------------------------------------------
' Result シートの全データをクリアする (ヘッダー行は残す)
' ボタンに割り当てて使用する
'-----------------------------------------------------------------------------
Public Sub Result_Clear()
    Dim ws As Worksheet
    Set ws = GetOrCreateResultSheet()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Row

    If lastRow <= 1 Then
        MsgBox "クリアするデータがありません。", vbInformation
        Exit Sub
    End If

    If MsgBox("Result シートの全データ (" & (lastRow - 1) & " 件) を削除しますか？", _
              vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ws.Rows("2:" & lastRow).Delete
    MsgBox "Result シートをクリアしました。", vbInformation
End Sub

'-----------------------------------------------------------------------------
' Result シートをアクティブにしてジャンプする
'-----------------------------------------------------------------------------
Public Sub Result_Activate()
    Dim ws As Worksheet
    Set ws = GetOrCreateResultSheet()
    ws.Activate
    ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Offset(1, 0).Select
End Sub

'=============================================================================
' プライベート関数
'=============================================================================

'-----------------------------------------------------------------------------
' Result シートを取得または新規作成する
'-----------------------------------------------------------------------------
Private Function GetOrCreateResultSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_RESULT)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_RESULT
        Call InitResultSheet(ws)
    End If

    Set GetOrCreateResultSheet = ws
End Function

'-----------------------------------------------------------------------------
' Result シートの初期レイアウトを設定する
'-----------------------------------------------------------------------------
Private Sub InitResultSheet(ws As Worksheet)
    ws.sheet_view.ShowGridLines = False   ' なければコメントアウト

    ' ヘッダーラベル
    ws.Cells(1, COL_NO).Value       = "No."
    ws.Cells(1, COL_DATETIME).Value = "実行日時"
    ws.Cells(1, COL_DEVICE).Value   = "機器名"
    ws.Cells(1, COL_ADDRESS).Value  = "VISAアドレス"
    ws.Cells(1, COL_CONNTYPE).Value = "接続方式"
    ws.Cells(1, COL_COMMAND).Value  = "コマンド / アクション"
    ws.Cells(1, COL_RESPONSE).Value = "応答結果"
    ws.Cells(1, COL_STATUS).Value   = "ステータス"
    ws.Cells(1, COL_NOTES).Value    = "備考"

    ' ヘッダー書式
    With ws.Range(ws.Cells(1, COL_NO), ws.Cells(1, COL_NOTES))
        .Font.Name            = "Meiryo UI"
        .Font.Bold            = True
        .Font.Color           = RGB(255, 255, 255)
        .Interior.Color       = RGB(31, 78, 121)
        .HorizontalAlignment  = xlCenter
        .VerticalAlignment    = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    ' 列幅
    ws.Columns(COL_NO).ColumnWidth       = 6
    ws.Columns(COL_DATETIME).ColumnWidth = 20
    ws.Columns(COL_DEVICE).ColumnWidth   = 16
    ws.Columns(COL_ADDRESS).ColumnWidth  = 32
    ws.Columns(COL_CONNTYPE).ColumnWidth = 14
    ws.Columns(COL_COMMAND).ColumnWidth  = 24
    ws.Columns(COL_RESPONSE).ColumnWidth = 36
    ws.Columns(COL_STATUS).ColumnWidth   = 10
    ws.Columns(COL_NOTES).ColumnWidth    = 20

    ' 先頭行固定・オートフィルター
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Rows(1).AutoFilter
    ws.Cells(1, 1).Select
End Sub

'-----------------------------------------------------------------------------
' VISAアドレスから接続方式を判定する
'-----------------------------------------------------------------------------
Private Function GetConnectionType(address As String) As String
    Dim a As String: a = UCase(Trim(address))
    If InStr(a, "GPIB") > 0 Then
        GetConnectionType = "GPIB"
    ElseIf InStr(a, "::SOCKET") > 0 Then
        GetConnectionType = "LAN Socket"
    ElseIf InStr(a, "HISLIP") > 0 Then
        GetConnectionType = "LAN HiSLIP"
    ElseIf InStr(a, "TCPIP") > 0 Then
        GetConnectionType = "LAN VXI-11"
    Else
        GetConnectionType = "UNKNOWN"
    End If
End Function

'-----------------------------------------------------------------------------
' Config シートから機器名に対応する VISAアドレスを検索する
'-----------------------------------------------------------------------------
Private Function FindAddressByDevice(wsConfig As Worksheet, deviceName As String) As String
    Dim lastRow As Long
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(wsConfig.Cells(i, 1).Value) = deviceName Then
            FindAddressByDevice = Trim(wsConfig.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    FindAddressByDevice = ""
End Function
