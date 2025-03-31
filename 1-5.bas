'''---------------------------------------------------------
' 1. 処理高速化のためのApplication.ScreenUpdating/Calculation制御
'''---------------------------------------------------------
Sub 高速処理制御サンプル()
    ' 変数宣言
    Dim lng_開始時間 As Long
    Dim lng_終了時間 As Long
    Dim ws_対象 As Worksheet
    
    ' 処理時間計測開始
    lng_開始時間 = Timer
    
    ' 画面更新と計算を一時停止（処理高速化）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' ここに実際の処理コードを記述
    Set ws_対象 = ThisWorkbook.Worksheets("Sheet1")
    ' 大量のデータ処理、集計などを行う...
    
NormalExit:
    ' 画面更新と計算を元に戻す（必ず実行される必要がある）
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
    ' 処理時間計測終了
    lng_終了時間 = Timer
    MsgBox "処理が完了しました。処理時間: " & Format(lng_終了時間 - lng_開始時間, "0.00") & "秒"
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Resume NormalExit
End Sub

'''---------------------------------------------------------
' 2. 頻出エラーハンドリングのテンプレート（On Error構文）
'''---------------------------------------------------------
Sub エラーハンドリングサンプル()
    ' 変数宣言
    Dim str_ファイルパス As String
    Dim wb_対象 As Workbook
    
    On Error GoTo ErrorHandler
    
    ' ファイルパスの設定
    str_ファイルパス = ThisWorkbook.Path & "\データ.xlsx"
    
    ' ファイルの存在チェック
    If Dir(str_ファイルパス) = "" Then
        MsgBox "ファイルが見つかりません: " & str_ファイルパス, vbExclamation
        Exit Sub
    End If
    
    ' ブックを開く
    Set wb_対象 = Workbooks.Open(str_ファイルパス)
    
    ' 何らかの処理...
    
    ' 正常終了時の処理
    wb_対象.Close SaveChanges:=True
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    ' エラー番号ごとの対応
    Select Case Err.Number
        Case 1004 ' アプリケーションまたはオブジェクト定義のエラー
            MsgBox "ファイルの書式に問題があります: " & Err.Description, vbCritical
        Case 9 ' 添え字が有効範囲にありません
            MsgBox "配列のインデックスが範囲外です: " & Err.Description, vbCritical
        Case 91 ' オブジェクト変数または With ブロック変数が設定されていません
            MsgBox "オブジェクトが正しく設定されていません: " & Err.Description, vbCritical
        Case Else
            MsgBox "予期せぬエラーが発生しました: " & vbCrLf & _
                  "エラー番号: " & Err.Number & vbCrLf & _
                  "説明: " & Err.Description, vbCritical
    End Select
    
    ' クリーンアップ処理
    On Error Resume Next ' エラーを無視して後処理を実行
    If Not wb_対象 Is Nothing Then
        wb_対象.Close SaveChanges:=False
    End If
    
    ' エラーログを記録（オプション）
    Call エラーログ記録(Err.Number, Err.Description, "エラーハンドリングサンプル")
End Sub

' エラーログを記録する補助関数
Private Sub エラーログ記録(lng_エラー番号 As Long, str_エラー内容 As String, str_プロシージャ名 As String)
    Dim str_ログパス As String
    Dim int_ファイル番号 As Integer
    
    On Error Resume Next
    
    str_ログパス = ThisWorkbook.Path & "\error_log.txt"
    int_ファイル番号 = FreeFile
    
    Open str_ログパス For Append As #int_ファイル番号
    Print #int_ファイル番号, Format(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
                           "プロシージャ: " & str_プロシージャ名 & vbTab & _
                           "エラー番号: " & lng_エラー番号 & vbTab & _
                           "内容: " & str_エラー内容
    Close #int_ファイル番号
End Sub

'''---------------------------------------------------------
' 3. セル範囲を配列に格納して高速処理する方法
'''---------------------------------------------------------
Sub 配列処理高速化サンプル()
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim rng_対象範囲 As Range
    Dim var_データ配列 As Variant
    Dim lng_行数 As Long
    Dim lng_列数 As Long
    Dim lng_行 As Long
    Dim lng_列 As Long
    Dim lng_合計値 As Long
    Dim lng_開始時間 As Long
    Dim lng_終了時間 As Long
    
    ' 処理時間計測開始
    lng_開始時間 = Timer
    
    ' 対象シートとデータ範囲の設定
    Set ws_データ = ThisWorkbook.Worksheets("データ")
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    Set rng_対象範囲 = ws_データ.Range("A1:C" & lng_最終行)
    
    ' セル範囲を一度に配列に読み込む（これが高速化の鍵）
    var_データ配列 = rng_対象範囲.Value
    
    ' 配列の次元を取得
    lng_行数 = UBound(var_データ配列, 1)
    lng_列数 = UBound(var_データ配列, 2)
    
    ' 配列を使った処理例（C列の数値合計）
    lng_合計値 = 0
    For lng_行 = 1 To lng_行数
        ' 数値かどうかをチェック
        If IsNumeric(var_データ配列(lng_行, 3)) Then
            lng_合計値 = lng_合計値 + CLng(var_データ配列(lng_行, 3))
        End If
    Next lng_行
    
    ' 処理結果を別の範囲に一度に書き込む例
    Dim var_結果配列() As Variant
    ReDim var_結果配列(1 To lng_行数, 1 To 2)
    
    For lng_行 = 1 To lng_行数
        var_結果配列(lng_行, 1) = var_データ配列(lng_行, 1) ' A列の値
        ' 何らかの計算を行う
        var_結果配列(lng_行, 2) = "処理済: " & var_データ配列(lng_行, 2) ' B列の値を加工
    Next lng_行
    
    ' 結果を一度にシートに書き込む
    ws_データ.Range("E1:F" & lng_行数).Value = var_結果配列
    
    ' 処理時間計測終了
    lng_終了時間 = Timer
    MsgBox "処理が完了しました。" & vbCrLf & _
           "合計値: " & lng_合計値 & vbCrLf & _
           "処理時間: " & Format(lng_終了時間 - lng_開始時間, "0.00") & "秒", vbInformation
End Sub

'''---------------------------------------------------------
' 4. シート間/ブック間のデータコピー効率化
'''---------------------------------------------------------
Sub シート間データコピー効率化()
    ' 変数宣言
    Dim wb_元ブック As Workbook
    Dim wb_先ブック As Workbook
    Dim ws_元シート As Worksheet
    Dim ws_先シート As Worksheet
    Dim rng_コピー元 As Range
    Dim rng_コピー先 As Range
    Dim lng_最終行 As Long
    Dim lng_最終列 As Long
    Dim var_データ As Variant
    
    ' 画面更新を停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' 元ブックと先ブックを設定
    Set wb_元ブック = ThisWorkbook
    
    ' 先ブックを開く、または新規作成
    On Error Resume Next
    Set wb_先ブック = Workbooks.Open(ThisWorkbook.Path & "\出力先.xlsx")
    If wb_先ブック Is Nothing Then
        Set wb_先ブック = Workbooks.Add
        wb_先ブック.SaveAs ThisWorkbook.Path & "\出力先.xlsx"
    End If
    On Error GoTo ErrorHandler
    
    ' シートの設定
    Set ws_元シート = wb_元ブック.Worksheets("元データ")
    
    ' 先ブックに同名シートがあるか確認し、なければ作成
    On Error Resume Next
    Set ws_先シート = wb_先ブック.Worksheets("コピー先")
    If ws_先シート Is Nothing Then
        Set ws_先シート = wb_先ブック.Worksheets.Add
        ws_先シート.Name = "コピー先"
    End If
    On Error GoTo ErrorHandler
    
    ' コピー元範囲の取得
    lng_最終行 = ws_元シート.Cells(ws_元シート.Rows.Count, "A").End(xlUp).Row
    lng_最終列 = ws_元シート.Cells(1, ws_元シート.Columns.Count).End(xlToLeft).Column
    Set rng_コピー元 = ws_元シート.Range(ws_元シート.Cells(1, 1), ws_元シート.Cells(lng_最終行, lng_最終列))
    
    ' データを配列に格納（効率化のポイント）
    var_データ = rng_コピー元.Value
    
    ' コピー先のシートをクリア
    ws_先シート.Cells.Clear
    
    ' コピー先範囲を設定して一括書き込み
    Set rng_コピー先 = ws_先シート.Range(ws_先シート.Cells(1, 1), ws_先シート.Cells(lng_最終行, lng_最終列))
    rng_コピー先.Value = var_データ
    
    ' 書式のコピー（オプション）
    rng_コピー元.Copy
    rng_コピー先.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' 先ブックを保存
    wb_先ブック.Save
    
NormalExit:
    ' 設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "データのコピーが完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume NormalExit
End Sub

'''---------------------------------------------------------
' 5. マクロ実行時間の計測と最適化手法
'''---------------------------------------------------------
Sub マクロ実行時間計測()
    ' 変数宣言
    Dim dbl_開始時間 As Double
    Dim dbl_終了時間 As Double
    Dim str_結果 As String
    Dim i As Long
    
    ' 開始時間を記録
    dbl_開始時間 = Timer
    
    ' パフォーマンス設定
    Call パフォーマンス設定(True)
    
    ' ここに処理コードを記述
    ' パフォーマンス測定のデモ
    For i = 1 To 10000
        ' 何か処理をする（例：シートに値を書き込む）
        ThisWorkbook.Worksheets(1).Cells(1, 1).Value = i
    Next i
    
    ' パフォーマンス設定を元に戻す
    Call パフォーマンス設定(False)
    
    ' 終了時間を記録
    dbl_終了時間 = Timer
    
    ' 実行時間を計算
    str_結果 = "処理時間: " & Format(dbl_終了時間 - dbl_開始時間, "0.00") & " 秒"
    Debug.Print str_結果
    MsgBox str_結果
End Sub

' パフォーマンス設定を一括で行う関数
Private Sub パフォーマンス設定(bln_高速化モード As Boolean)
    With Application
        If bln_高速化モード Then
            ' 高速化モード
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
            .DisplayAlerts = False
        Else
            ' 通常モード（元に戻す）
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
            .DisplayAlerts = True
        End If
    End With
End Sub

' パフォーマンス比較のためのテスト関数
Sub パフォーマンス比較テスト()
    Dim dbl_時間1 As Double
    Dim dbl_時間2 As Double
    
    ' 方法1: セルごとに書き込む（低速）
    dbl_時間1 = Timer
    Call 方法1_セルごと書き込み
    dbl_時間1 = Timer - dbl_時間1
    
    ' シートをクリア
    Worksheets(1).Range("A:B").Clear
    
    ' 方法2: 配列を使用して一括書き込み（高速）
    dbl_時間2 = Timer
    Call 方法2_配列で一括書き込み
    dbl_時間2 = Timer - dbl_時間2
    
    ' 結果表示
    MsgBox "方法1(セルごと): " & Format(dbl_時間1, "0.00") & " 秒" & vbCrLf & _
           "方法2(配列一括): " & Format(dbl_時間2, "0.00") & " 秒" & vbCrLf & _
           "高速化率: " & Format(dbl_時間1 / dbl_時間2, "0.00") & " 倍"
End Sub

' 方法1: セルごとに書き込む（低速）
Private Sub 方法1_セルごと書き込み()
    Dim i As Long
    Dim ws As Worksheet
    
    Set ws = Worksheets(1)
    
    For i = 1 To 10000
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = "テスト" & i
    Next i
End Sub

' 方法2: 配列を使用して一括書き込み（高速）
Private Sub 方法2_配列で一括書き込み()
    Dim i As Long
    Dim ws As Worksheet
    Dim arr_データ() As Variant
    
    Set ws = Worksheets(1)
    
    ' 配列の初期化
    ReDim arr_データ(1 To 10000, 1 To 2)
    
    ' 配列にデータを格納
    For i = 1 To 10000
        arr_データ(i, 1) = i
        arr_データ(i, 2) = "テスト" & i
    Next i
    
    ' 一度にシートに書き込み
    ws.Range("A1:B10000").Value = arr_データ
End Sub