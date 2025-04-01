MsgBox "言語設定を変更するには、以下の手順で操作してください:" & vbCrLf & vbCrLf & _
           "1. [ファイル] メニュー → [オプション] を選択" & vbCrLf & _
           "2. [言語] を選択" & vbCrLf & _
           "3. [Office の表示言語] で日本語を選択" & vbCrLf & _
           "4. [OK] をクリックし、Excelを再起動", _
           vbInformation, "言語設定変更"
    
    ' この関数では実際に言語を変更することはできません
    ' （Excelの言語設定はVBAから直接変更できない）
End Sub

'''---------------------------------------------------------
' 言語を英語に切り替える（ボタン用）
'''---------------------------------------------------------
Public Sub 言語切替_英語()
    ' ユーザーに対して言語切替の説明
    MsgBox "To change the language settings, follow these steps:" & vbCrLf & vbCrLf & _
           "1. Click [File] menu → [Options]" & vbCrLf & _
           "2. Select [Language]" & vbCrLf & _
           "3. Set English as the [Display Language]" & vbCrLf & _
           "4. Click [OK] and restart Excel", _
           vbInformation, "Language Settings"
    
    ' この関数では実際に言語を変更することはできません
    ' （Excelの言語設定はVBAから直接変更できない）
End Sub

'''---------------------------------------------------------
' 5. 大規模アプリケーションのパフォーマンス最適化
'''---------------------------------------------------------
' ***************************************************************************************************
' モジュール: パフォーマンス最適化
' 機能: 大規模アプリケーションのパフォーマンスを最適化するための技術
' ***************************************************************************************************
Option Explicit

' パフォーマンス計測用変数
Private dbl_開始時間 As Double
Private dbl_経過時間 As Double
Private arr_測定ポイント() As Variant
Private lng_測定数 As Long

'''---------------------------------------------------------
' パフォーマンス最適化のメイン処理
'''---------------------------------------------------------
Sub パフォーマンス最適化実装()
    ' 変数宣言
    Dim ws_情報 As Worksheet
    Dim lng_行数 As Long
    Dim lng_データ件数 As Long
    Dim bln_大規模データフラグ As Boolean
    Dim i As Long
    
    ' 測定ポイントの初期化
    ReDim arr_測定ポイント(1 To 10, 1 To 3) ' ポイント名、時間、メモリ使用量
    lng_測定数 = 0
    
    ' 計測開始
    パフォーマンス測定開始
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("パフォーマンス情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "パフォーマンス情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo 0
    
    ' タイトル設定
    ws_情報.Range("A1").Value = "大規模アプリケーションのパフォーマンス最適化"
    ws_情報.Range("A1").Font.Size = 14
    ws_情報.Range("A1").Font.Bold = True
    
    ' 測定ポイントを記録
    パフォーマンス測定ポイント記録 "初期化完了"
    
    ' データ規模の判定（行数に基づく）
    lng_行数 = Application.WorksheetFunction.Min(100000, Application.Rows.Count - 10)
    lng_データ件数 = 10000 ' デモ用に10000件に制限
    
    bln_大規模データフラグ = (lng_データ件数 > 5000)
    
    ws_情報.Range("A3").Value = "テストデータサイズ:"
    ws_情報.Range("B3").Value = Format(lng_データ件数, "#,##0") & " 行"
    
    ws_情報.Range("A4").Value = "大規模データフラグ:"
    ws_情報.Range("B4").Value = IIf(bln_大規模データフラグ, "ON（最適化モード）", "OFF（通常モード）")
    
    ' パフォーマンス比較（通常 vs 最適化）のデモ
    Dim dbl_通常処理時間 As Double
    Dim dbl_最適化処理時間 As Double
    
    ' サンプルデータの生成
    Application.StatusBar = "サンプルデータを生成中..."
    Call サンプルデータ生成(ws_情報, lng_データ件数, bln_大規模データフラグ)
    
    ' 測定ポイントを記録
    パフォーマンス測定ポイント記録 "データ生成完了"
    
    ' 通常処理時間の計測
    Application.StatusBar = "通常処理を実行中..."
    dbl_通常処理時間 = 処理時間計測_通常(ws_情報, lng_データ件数)
    
    ' 測定ポイントを記録
    パフォーマンス測定ポイント記録 "通常処理完了"
    
    ' 最適化処理時間の計測
    Application.StatusBar = "最適化処理を実行中..."
    dbl_最適化処理時間 = 処理時間計測_最適化(ws_情報, lng_データ件数)
    
    ' 測定ポイントを記録
    パフォーマンス測定ポイント記録 "最適化処理完了"
    
    ' 性能比較結果の表示
    ws_情報.Range("A6").Value = "性能比較結果:"
    ws_情報.Range("A6").Font.Bold = True
    
    ws_情報.Range("A7").Value = "処理方法"
    ws_情報.Range("B7").Value = "処理時間(秒)"
    ws_情報.Range("C7").Value = "比率(%)"
    ws_情報.Range("A7:C7").Font.Bold = True
    
    ' 通常処理の結果
    ws_情報.Range("A8").Value = "通常処理"
    ws_情報.Range("B8").Value = Format(dbl_通常処理時間, "0.000")
    ws_情報.Range("C8").Value = "100%"
    
    ' 最適化処理の結果
    ws_情報.Range("A9").Value = "最適化処理"
    ws_情報.Range("B9").Value = Format(dbl_最適化処理時間, "0.000")
    
    ' 速度向上比率の計算
    If dbl_通常処理時間 > 0 Then
        ws_情報.Range("C9").Value = Format(dbl_最適化処理時間 / dbl_通常処理時間 * 100, "0.0") & "%"
    Else
        ws_情報.Range("C9").Value = "N/A"
    End If
    
    ' 書式設定
    ws_情報.Range("A8:C8").Interior.Color = RGB(255, 235, 220) ' 薄橙（低速）
    ws_情報.Range("A9:C9").Interior.Color = RGB(220, 255, 220) ' 薄緑（高速）
    
    ' 測定ポイントデータの表示
    ws_情報.Range("A11").Value = "パフォーマンス測定ポイント:"
    ws_情報.Range("A11").Font.Bold = True
    
    ws_情報.Range("A12").Value = "測定ポイント"
    ws_情報.Range("B12").Value = "経過時間(秒)"
    ws_情報.Range("C12").Value = "累積時間(秒)"
    ws_情報.Range("A12:C12").Font.Bold = True
    
    ' 測定ポイントのデータを表示
    For i = 1 To lng_測定数
        ws_情報.Range("A" & (12 + i)).Value = arr_測定ポイント(i, 1)
        ws_情報.Range("B" & (12 + i)).Value = Format(arr_測定ポイント(i, 2), "0.000")
        ws_情報.Range("C" & (12 + i)).Value = Format(arr_測定ポイント(i, 3), "0.000")
    Next i
    
    ' パフォーマンス最適化のベストプラクティス
    ws_情報.Range("A" & (14 + lng_測定数)).Value = "パフォーマンス最適化ベストプラクティス:"
    ws_情報.Range("A" & (14 + lng_測定数)).Font.Bold = True
    
    Dim arr_ベストプラクティス As Variant
    arr_ベストプラクティス = Array( _
        "1. スクリーン更新・計算の一時停止: Application.ScreenUpdating = False", _
        "2. 配列によるデータの一括処理: セルごとではなく配列で処理", _
        "3. シート・ブック操作の最小化: 無駄なシート間の移動を避ける", _
        "4. 数式ではなく値のコピー: .Value = .Value または .PasteSpecial xlPasteValues", _
        "5. 範囲指定の最適化: Cells(i, j) よりも Range(...).Cells(i, j) が高速", _
        "6. 条件付き書式やフィルタの適用回数制限: まとめて一度だけ適用", _
        "7. 大規模データでのAutoFilterよりAdvancedFilterの使用", _
        "8. メモリリークの防止: オブジェクト変数の解放 (Set obj = Nothing)", _
        "9. 適切なデータ型の使用: Variantを避け、最適な型を選択", _
        "10. コード内のループ最適化: 不要な繰り返し計算の排除", _
        "11. イベントの一時停止: Application.EnableEvents = False", _
        "12. バッチ処理の導入: 一度に全てではなく処理を分割", _
        "13. パフォーマンス計測ポイントの設置: ボトルネック特定" _
    )
    
    For i = 0 To UBound(arr_ベストプラクティス)
        ws_情報.Range("A" & (15 + lng_測定数 + i)).Value = arr_ベストプラクティス(i)
    Next i
    
    ' パフォーマンスに関連するグラフの作成（オプション）
    Call パフォーマンスグラフ作成(ws_情報, lng_測定数)
    
    ' 書式設定
    ws_情報.Columns("A:C").AutoFit
    ws_情報.Range("A7:C9").Borders.LineStyle = xlContinuous
    ws_情報.Range("A12:C" & (12 + lng_測定数)).Borders.LineStyle = xlContinuous
    
    ' アクティブ化
    ws_情報.Activate
    ws_情報.Range("A1").Select
    
    ' 測定終了
    パフォーマンス測定終了
    
    ' パフォーマンス計測結果を表示
    dbl_経過時間 = Round(Timer - dbl_開始時間, 3)
    
    Application.StatusBar = False
    MsgBox "パフォーマンス測定が完了しました。" & vbCrLf & _
           "総実行時間: " & Format(dbl_経過時間, "0.000") & " 秒" & vbCrLf & _
           "性能向上率: " & Format((1 - dbl_最適化処理時間 / dbl_通常処理時間) * 100, "0.0") & "%", _
           vbInformation, "パフォーマンス最適化"
End Sub

'''---------------------------------------------------------
' パフォーマンス測定を開始する
'''---------------------------------------------------------
Private Sub パフォーマンス測定開始()
    ' 開始時間を記録
    dbl_開始時間 = Timer
    
    ' アプリケーション設定の最適化（測定の正確性のため）
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayStatusBar = True
    End With
End Sub

'''---------------------------------------------------------
' パフォーマンス測定ポイントを記録する
' 引数:
'   Str_ポイント名 - 測定ポイントの名前
'''---------------------------------------------------------
Private Sub パフォーマンス測定ポイント記録(ByVal Str_ポイント名 As String)
    ' 測定数をインクリメント
    lng_測定数 = lng_測定数 + 1
    
    ' 配列の範囲を確認し、必要に応じて拡張
    If lng_測定数 > UBound(arr_測定ポイント, 1) Then
        ReDim Preserve arr_測定ポイント(1 To lng_測定数 + 5, 1 To 3)
    End If
    
    ' 測定データを記録
    arr_測定ポイント(lng_測定数, 1) = Str_ポイント名
    
    ' 直前のポイントからの経過時間
    If lng_測定数 = 1 Then
        arr_測定ポイント(lng_測定数, 2) = Timer - dbl_開始時間
    Else
        Dim dbl_前回時間 As Double
        dbl_前回時間 = arr_測定ポイント(lng_測定数 - 1, 3)
        arr_測定ポイント(lng_測定数, 2) = Timer - dbl_開始時間 - dbl_前回時間
    End If
    
    ' 開始からの累積時間
    arr_測定ポイント(lng_測定数, 3) = Timer - dbl_開始時間
End Sub

'''---------------------------------------------------------
' パフォーマンス測定を終了する
'''---------------------------------------------------------
Private Sub パフォーマンス測定終了()
    ' アプリケーション設定を元に戻す
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
    
    ' 総経過時間を記録
    dbl_経過時間 = Timer - dbl_開始時間
End Sub

'''---------------------------------------------------------
' サンプルデータを生成する関数
' 引数:
'   Ws_データ - データを書き込むワークシート
'   Lng_データ件数 - 生成するデータの行数
'   Bln_最適化フラグ - 最適化モードフラグ
'''---------------------------------------------------------
Private Sub サンプルデータ生成(ByVal Ws_データ As Worksheet, _
                       ByVal Lng_データ件数 As Long, _
                       ByVal Bln_最適化フラグ As Boolean)
    ' 変数宣言
    Dim lng_行 As Long
    Dim lng_列 As Long
    Dim arr_データ() As Variant
    
    ' 既存データをクリア
    Ws_データ.Range("F1:J" & Lng_データ件数 + 1).ClearContents
    
    ' ヘッダー行の設定
    Ws_データ.Range("F1").Value = "顧客ID"
    Ws_データ.Range("G1").Value = "契約日"
    Ws_データ.Range("H1").Value = "契約種別"
    Ws_データ.Range("I1").Value = "契約電力(kW)"
    Ws_データ.Range("J1").Value = "基本料金"
    
    ' 大規模データの場合は配列を使用（最適化モード）
    If Bln_最適化フラグ Then
        ' 配列の初期化
        ReDim arr_データ(1 To Lng_データ件数, 1 To 5)
        
        ' 配列にデータを格納
        For lng_行 = 1 To Lng_データ件数
            ' 進捗表示（100行ごと）
            If lng_行 Mod 100 = 0 Then
                Application.StatusBar = "サンプルデータ生成中... " & Format(lng_行 / Lng_データ件数, "0%")
                DoEvents
            End If
            
            ' データの生成
            arr_データ(lng_行, 1) = "C" & Format(lng_行, "000000")
            arr_データ(lng_行, 2) = DateAdd("d", -Int(Rnd * 1000), Date)
            
            Select Case Int(Rnd * 3)
                Case 0: arr_データ(lng_行, 3) = "従量電灯B"
                Case 1: arr_データ(lng_行, 3) = "従量電灯C"
                Case 2: arr_データ(lng_行, 3) = "低圧電力"
            End Select
            
            arr_データ(lng_行, 4) = Int(Rnd * 100) + 10
            arr_データ(lng_行, 5) = arr_データ(lng_行, 4) * (Int(Rnd * 10) + 15) * 10
        Next lng_行
        
        ' データを一括で書き込み
        Ws_データ.Range("F2").Resize(Lng_データ件数, 5).Value = arr_データ
    Else
        ' 通常モード（セルごとに書き込み）
        For lng_行 = 1 To Lng_データ件数
            ' 進捗表示（10行ごと）
            If lng_行 Mod 10 = 0 Then
                Application.StatusBar = "サンプルデータ生成中... " & Format(lng_行 / Lng_データ件数, "0%")
                DoEvents
            End If
            
            ' データの生成とセルへの書き込み
            Ws_データ.Cells(lng_行 + 1, 6).Value = "C" & Format(lng_行, "000000")
            Ws_データ.Cells(lng_行 + 1, 7).Value = DateAdd("d", -Int(Rnd * 1000), Date)
            
            Select Case Int(Rnd * 3)
                Case 0: Ws_データ.Cells(lng_行 + 1, 8).Value = "従量電灯B"
                Case 1: Ws_データ.Cells(lng_行 + 1, 8).Value = "従量電灯C"
                Case 2: Ws_データ.Cells(lng_行 + 1, 8).Value = "低圧電力"
            End Select
            
            Ws_データ.Cells(lng_行 + 1, 9).Value = Int(Rnd * 100) + 10
            Ws_データ.Cells(lng_行 + 1, 10).Value = Ws_データ.Cells(lng_行 + 1, 9).Value * (Int(Rnd * 10) + 15) * 10
        Next lng_行
    End If
    
    ' 書式設定
    Ws_データ.Range("F1:J1").Font.Bold = True
    Ws_データ.Range("G2:G" & Lng_データ件数 + 1).NumberFormat = "yyyy/mm/dd"
    Ws_データ.Range("I2:I" & Lng_データ件数 + 1).NumberFormat = "#,##0"
    Ws_データ.Range("J2:J" & Lng_データ件数 + 1).NumberFormat = "#,##0"
    
    ' 列幅の自動調整
    Ws_データ.Columns("F:J").AutoFit
End Sub

'''---------------------------------------------------------
' 通常処理の実行時間を計測する関数
' 引数:
'   Ws_データ - データを処理するワークシート
'   Lng_データ件数 - 処理するデータの行数
' 戻り値:
'   Double - 処理にかかった時間（秒）
'''---------------------------------------------------------
Private Function 処理時間計測_通常(ByVal Ws_データ As Worksheet, _
                             ByVal Lng_データ件数 As Long) As Double
    ' 変数宣言
    Dim dbl_開始 As Double
    Dim lng_行 As Long
    Dim rng_契約電力 As Range
    Dim rng_基本料金 As Range
    Dim dbl_合計電力 As Double
    Dim dbl_合計料金 As Double
    Dim rng_対象セル As Range
    
    ' 計測開始
    dbl_開始 = Timer
    
    ' 通常の処理（セルごとの参照による集計）
    dbl_合計電力 = 0
    dbl_合計料金 = 0
    
    For lng_行 = 2 To Lng_データ件数 + 1
        ' 進捗表示（50行ごと）
        If lng_行 Mod 50 = 0 Then
            Application.StatusBar = "通常処理実行中... " & Format((lng_行 - 2) / Lng_データ件数, "0%")
            DoEvents
        End If
        
        ' 契約種別をチェック
        If Ws_データ.Cells(lng_行, 8).Value = "従量電灯B" Then
            ' 契約電力を加算
            dbl_合計電力 = dbl_合計電力 + Ws_データ.Cells(lng_行, 9).Value
            
            ' 基本料金を加算
            dbl_合計料金 = dbl_合計料金 + Ws_データ.Cells(lng_行, 10).Value
        End If
    Next lng_行
    
    ' 平均値の計算
    Dim dbl_平均電力 As Double
    Dim dbl_平均料金 As Double
    
    If dbl_合計電力 > 0 Then
        dbl_平均電力 = dbl_合計電力 / Lng_データ件数
    Else
        dbl_平均電力 = 0
    End If
    
    If dbl_合計料金 > 0 Then
        dbl_平均料金 = dbl_合計料金 / Lng_データ件数
    Else
        dbl_平均料金 = 0
    End If
    
    ' 書式設定の適用（通常の方法）
    Set rng_契約電力 = Ws_データ.Range("I2:I" & Lng_データ件数 + 1)
    Set rng_基本料金 = Ws_データ.Range("J2:J" & Lng_データ件数 + 1)
    
    ' 条件付き書式（100kW以上の契約電力を強調）
    rng_契約電力.FormatConditions.Delete
    
    Set rng_対象セル = rng_契約電力
    rng_対象セル.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="100"
    rng_対象セル.FormatConditions(1).Interior.Color = RGB(255, 200, 200)
    
    ' 結果をセルに表示
    Ws_データ.Range("F" & Lng_データ件数 + 3).Value = "合計契約電力:"
    Ws_データ.Range("G" & Lng_データ件数 + 3).Value = dbl_合計電力
    
    Ws_データ.Range("F" & Lng_データ件数 + 4).Value = "合計基本料金:"
    Ws_データ.Range("G" & Lng_データ件数 + 4).Value = dbl_合計料金
    
    ' 処理時間を返す
    処理時間計測_通常 = Timer - dbl_開始
End Function

'''---------------------------------------------------------
' 最適化処理の実行時間を計測する関数
' 引数:
'   Ws_データ - データを処理するワークシート
'   Lng_データ件数 - 処理するデータの行数
' 戻り値:
'   Double - 処理にかかった時間（秒）
'''---------------------------------------------------------
Private Function 処理時間計測_最適化(ByVal Ws_データ As Worksheet, _
                              ByVal Lng_データ件数 As Long) As Double
    ' 変数宣言
    Dim dbl_開始 As Double
    Dim lng_行 As Long
    Dim arr_データ As Variant
    Dim dbl_合計電力 As Double
    Dim dbl_合計料金 As Double
    Dim arr_条件付き書式対象() As Long
    Dim lng_対象数 As Long
    
    ' 計測開始
    dbl_開始 = Timer
    
    ' 最適化処理（配列を使用した一括処理）
    ' データを一括で配列に取得
    arr_データ = Ws_データ.Range("F2:J" & Lng_データ件数 + 1).Value
    
    ' 配列の初期化
    ReDim arr_条件付き書式対象(1 To Lng_データ件数)
    lng_対象数 = 0
    
    ' 配列内でデータを処理
    dbl_合計電力 = 0
    dbl_合計料金 = 0
    
    For lng_行 = 1 To Lng_データ件数
        ' 進捗表示（1000行ごと）
        If lng_行 Mod 1000 = 0 Then
            Application.StatusBar = "最適化処理実行中... " & Format(lng_行 / Lng_データ件数, "0%")
            DoEvents
        End If
        
        ' 契約種別をチェック
        If arr_データ(lng_行, 3) = "従量電灯B" Then
            ' 契約電力を加算
            dbl_合計電力 = dbl_合計電力 + arr_データ(lng_行, 4)
            
            ' 基本料金を加算
            dbl_合計料金 = dbl_合計料金 + arr_データ(lng_行, 5)
        End If
        
        ' 条件付き書式用のインデックスを記録（100kW以上）
        If arr_データ(lng_行, 4) > 100 Then
            lng_対象数 = lng_対象数 + 1
            arr_条件付き書式対象(lng_対象数) = lng_行
        End If
    Next lng_行
    
    ' 平均値の計算
    Dim dbl_平均電力 As Double
    Dim dbl_平均料金 As Double
    
    If dbl_合計電力 > 0 Then
        dbl_平均電力 = dbl_合計電力 / Lng_データ件数
    Else
        dbl_平均電力 = 0
    End If
    
    If dbl_合計料金 > 0 Then
        dbl_平均料金 = dbl_合計料金 / Lng_データ件数
    Else
        dbl_平均料金 = 0
    End If
    
    ' 条件付き書式の適用（最適化版）
    Ws_データ.Range("I2:I" & Lng_データ件数 + 1).FormatConditions.Delete
    
    ' 条件に一致するセルにのみ書式を適用
    If lng_対象数 > 0 Then
        Dim rng_対象セル As Range
        Set rng_対象セル = Ws_データ.Cells(1 + arr_条件付き書式対象(1), 9)
        
        ' 対象セルの結合
        For lng_行 = 2 To lng_対象数
            Set rng_対象セル = Union(rng_対象セル, Ws_データ.Cells(1 + arr_条件付き書式対象(lng_行), 9))
        Next lng_行
        
        ' 一括で書式設定
        rng_対象セル.Interior.Color = RGB(255, 200, 200)
    End If
    
    ' 結果をセルに表示
    Ws_データ.Range("H" & Lng_データ件数 + 3).Value = "合計契約電力:"
    Ws_データ.Range("I" & Lng_データ件数 + 3).Value = dbl_合計電力
    
    Ws_データ.Range("H" & Lng_データ件数 + 4).Value = "合計基本料金:"
    Ws_データ.Range("I" & Lng_データ件数 + 4).Value = dbl_合計料金
    
    ' 処理時間を返す
    処理時間計測_最適化 = Timer - dbl_開始
End Function

'''---------------------------------------------------------
' パフォーマンスグラフを作成する関数
' 引数:
'   Ws_グラフ - グラフを作成するワークシート
'   Lng_測定数 - 測定ポイントの数
'''---------------------------------------------------------
Private Sub パフォーマンスグラフ作成(ByVal Ws_グラフ As Worksheet, ByVal Lng_測定数 As Long)
    ' グラフが既に存在する場合は削除
    On Error Resume Next
    Dim cht_既存 As ChartObject
    For Each cht_既存 In Ws_グラフ.ChartObjects
        cht_既存.Delete
    Next cht_既存
    On Error GoTo 0
    
    ' 測定ポイントがない場合は終了
    If Lng_測定数 <= 0 Then Exit Sub
    
    ' グラフオブジェクトの作成
    Dim cht_パフォーマンス As ChartObject
    Set cht_パフォーマンス = Ws_グラフ.ChartObjects.Add(Left:=400, Top:=50, Width:=450, Height:=250)
    
    ' グラフの設定
    With cht_パフォーマンス.Chart
        ' グラフ種類の設定
        .ChartType = xlColumnClustered
        
        ' データ範囲の設定
        .SetSourceData Source:=Ws_グラフ.Range("A13:B" & (12 + Lng_測定数))
        
        ' タイトルの設定
        .HasTitle = True
        .ChartTitle.Text = "処理ポイント別の所要時間"
        
        ' 軸ラベルの設定
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "処理ポイント"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "所要時間 (秒)"
        
        ' 凡例を非表示
        .HasLegend = False
        
        ' グラフエリアの書式設定
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(240, 240, 240)
        
        ' プロットエリアの書式設定
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' データラベルの追加
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.ShowValue = True
        .SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font.Size = 9
    End With
End Sub    ws_情報.Range("A4").Value = "ビルド番号:"
    ws_情報.Range("B4").Value = Application.Build
    
    ws_情報.Range("A5").Value = "対応状況:"
    
    If bln_対応バージョン Then
        ws_情報.Range("B5").Value = "対応バージョン"
        ws_情報.Range("B5").Interior.Color = RGB(220, 255, 220) ' 薄緑
    Else
        ws_情報.Range("B5").Value = "非対応バージョン（" & STR_最小対応バージョン & "以上必要）"
        ws_情報.Range("B5").Interior.Color = RGB(255, 220, 220) ' 薄赤
    End If
    
    ' バージョン別の機能対応表
    ws_情報.Range("A7").Value = "バージョン別機能対応:"
    ws_情報.Range("A7").Font.Bold = True
    
    ws_情報.Range("A8").Value = "機能"
    ws_情報.Range("B8").Value = "Excel 2007"
    ws_情報.Range("C8").Value = "Excel 2010"
    ws_情報.Range("D8").Value = "Excel 2013"
    ws_情報.Range("E8").Value = "Excel 2016/365"
    ws_情報.Range("A8:E8").Font.Bold = True
    
    ' 機能対応データ
    Dim arr_機能 As Variant
    arr_機能 = Array( _
        Array("リボンUIカスタマイズ", "△", "○", "○", "○"), _
        Array("Power Query", "×", "△", "○", "○"), _
        Array("PowerPivot", "×", "△", "○", "○"), _
        Array("スライサー", "×", "○", "○", "○"), _
        Array("フラッシュフィル", "×", "×", "○", "○"), _
        Array("新統計関数", "×", "×", "×", "○"), _
        Array("IFS/TEXTJOIN関数", "×", "×", "×", "○"), _
        Array("DDE通信", "○", "○", "○", "△"), _
        Array("VBA7 (64bit対応)", "×", "○", "○", "○"), _
        Array("新グラフ種類", "×", "×", "△", "○") _
    )
    
    ' 機能対応表の表示
    For i = 0 To UBound(arr_機能)
        ws_情報.Range("A" & (9 + i)).Value = arr_機能(i)(0) ' 機能名
        
        ' Excel 2007 (バージョン12.0)
        ws_情報.Range("B" & (9 + i)).Value = arr_機能(i)(1)
        
        ' Excel 2010 (バージョン14.0)
        ws_情報.Range("C" & (9 + i)).Value = arr_機能(i)(2)
        
        ' Excel 2013 (バージョン15.0)
        ws_情報.Range("D" & (9 + i)).Value = arr_機能(i)(3)
        
        ' Excel 2016/365 (バージョン16.0)
        ws_情報.Range("E" & (9 + i)).Value = arr_機能(i)(4)
        
        ' セル色の設定
        For j = 1 To 4
            Select Case arr_機能(i)(j)
                Case "○"
                    ws_情報.Cells(9 + i, j + 1).Interior.Color = RGB(220, 255, 220) ' 薄緑（対応）
                Case "△"
                    ws_情報.Cells(9 + i, j + 1).Interior.Color = RGB(255, 255, 200) ' 薄黄（一部対応）
                Case "×"
                    ws_情報.Cells(9 + i, j + 1).Interior.Color = RGB(255, 220, 220) ' 薄赤（非対応）
            End Select
        Next j
    Next i
    
    ' 現在のバージョンの列を強調表示
    Dim lng_バージョン列 As Long
    Select Case Int(dbl_Excelバージョン)
        Case 12: lng_バージョン列 = 2 ' Excel 2007
        Case 14: lng_バージョン列 = 3 ' Excel 2010
        Case 15: lng_バージョン列 = 4 ' Excel 2013
        Case 16: lng_バージョン列 = 5 ' Excel 2016/365
        Case Else: lng_バージョン列 = 0
    End Select
    
    If lng_バージョン列 > 0 Then
        ws_情報.Columns(lng_バージョン列).Interior.Color = RGB(240, 240, 255) ' 薄紫で列全体を強調
        ws_情報.Cells(8, lng_バージョン列).Interior.Color = RGB(200, 200, 255) ' 濃い紫でヘッダを強調
    End If
    
    ' 互換性対応のベストプラクティス
    Dim arr_ベストプラクティス As Variant
    arr_ベストプラクティス = Array( _
        "Excel互換性対応のベストプラクティス:", _
        "", _
        "1. 最低対応バージョンを明確に定義し、それ以前の機能に依存しない", _
        "2. バージョン依存の機能は条件分岐で回避策を用意", _
        "3. 新しい関数（IFS, TEXTJOIN等）は古いバージョンでは使えないことに注意", _
        "4. FileFormatプロパティでファイル形式を明示的に指定（.xlsx vs .xls）", _
        "5. マクロを記録して機能の違いを確認（特にグラフやピボットテーブル）", _
        "6. 複雑な機能ほどバージョン間の互換性に問題が発生しやすい", _
        "7. バージョン確認コードを含め、非対応の場合は明示的にエラーメッセージを表示" _
    )
    
    ws_情報.Range("A21").Value = arr_ベストプラクティス(0)
    ws_情報.Range("A21").Font.Bold = True
    
    For i = 1 To UBound(arr_ベストプラクティス)
        ws_情報.Range("A" & (21 + i)).Value = arr_ベストプラクティス(i)
    Next i
    
    ' バージョン検出コードサンプル
    ws_情報.Range("A31").Value = "バージョン検出コードサンプル:"
    ws_情報.Range("A31").Font.Bold = True
    
    Dim arr_コード例 As Variant
    arr_コード例 = Array( _
        "' バージョンに応じた処理を行う関数", _
        "Public Function バージョン対応処理() As Boolean", _
        "    Dim dbl_Version As Double", _
        "    dbl_Version = CDbl(Application.Version)", _
        "", _
        "    ' Excel 2010以降で利用可能な機能", _
        "    If dbl_Version >= 14 Then", _
        "        ' Excel 2010以降の処理", _
        "        ' スライサーなどの新機能を使用", _
        "    Else", _
        "        ' Excel 2007以前の代替処理", _
        "        ' 通常のフィルター等で代替", _
        "    End If", _
        "", _
        "    ' Excel 2013以降で利用可能な機能", _
        "    If dbl_Version >= 15 Then", _
        "        ' フラッシュフィル等の機能を使用", _
        "    End If", _
        "", _
        "    ' Excel 2016以降で利用可能な機能", _
        "    If dbl_Version >= 16 Then", _
        "        ' TEXTJOIN等の新関数を使用", _
        "    End If", _
        "", _
        "    Return True", _
        "End Function" _
    )
    
    For i = 0 To UBound(arr_コード例)
        ws_情報.Range("A" & (32 + i)).Value = arr_コード例(i)
    Next i
    
    ' 書式設定
    ws_情報.Columns("A:E").AutoFit
    ws_情報.Range("A8:E" & (9 + UBound(arr_機能))).Borders.LineStyle = xlContinuous
    
    ' アクティブ化
    ws_情報.Activate
    ws_情報.Range("A1").Select
End Sub

'''---------------------------------------------------------
' Excelのバージョン番号から製品名を取得する関数
'''---------------------------------------------------------
Private Function Excelバージョン名取得(ByVal Dbl_バージョン As Double) As String
    ' バージョン番号から製品名を判定
    Select Case Int(Dbl_バージョン)
        Case 8:  Excelバージョン名取得 = "Excel 97"
        Case 9:  Excelバージョン名取得 = "Excel 2000"
        Case 10: Excelバージョン名取得 = "Excel 2002/XP"
        Case 11: Excelバージョン名取得 = "Excel 2003"
        Case 12: Excelバージョン名取得 = "Excel 2007"
        Case 14: Excelバージョン名取得 = "Excel 2010"
        Case 15: Excelバージョン名取得 = "Excel 2013"
        Case 16
            ' Excel 2016とExcel 365はビルド番号で区別
            If Application.Build >= 13328 Then
                Excelバージョン名取得 = "Excel 365"
            Else
                Excelバージョン名取得 = "Excel 2016"
            End If
        Case Else: Excelバージョン名取得 = "不明なバージョン"
    End Select
End Function

'''---------------------------------------------------------
' バージョン依存機能の互換性確保のためのラッパー関数の例
'''---------------------------------------------------------
Private Function IFSラッパー(ParamArray 条件と値() As Variant) As Variant
    ' IFS関数のラッパー（Excel 2016以降でのみネイティブに利用可能）
    Dim i As Long
    
    ' Excelのバージョンをチェック
    If CDbl(Application.Version) >= 16 Then
        ' Excel 2016以降の場合はワークシート関数のIFSを使用
        Dim str_式 As String
        str_式 = "=IFS("
        
        For i = LBound(条件と値) To UBound(条件と値) Step 2
            ' 最後の要素の場合（デフォルト値）
            If i = UBound(条件と値) Then
                str_式 = str_式 & "TRUE," & 条件と値(i)
            Else
                str_式 = str_式 & 条件と値(i) & "," & 条件と値(i + 1)
                
                ' まだ条件と値のペアがあれば、カンマを追加
                If i + 2 <= UBound(条件と値) Then
                    str_式 = str_式 & ","
                End If
            End If
        Next i
        
        str_式 = str_式 & ")"
        
        ' 数式を評価
        On Error Resume Next
        IFSラッパー = Evaluate(str_式)
        If Err.Number <> 0 Then
            IFSラッパー = "エラー: " & Err.Description
        End If
        On Error GoTo 0
    Else
        ' Excel 2013以前の場合は自前でIFSの動作を実装
        For i = LBound(条件と値) To UBound(条件と値) Step 2
            ' 最後の要素の場合（デフォルト値）
            If i = UBound(条件と値) Then
                IFSラッパー = 条件と値(i)
                Exit Function
            End If
            
            ' 条件を評価
            On Error Resume Next
            Dim bln_条件 As Boolean
            bln_条件 = Evaluate(条件と値(i))
            
            If Err.Number <> 0 Then
                IFSラッパー = "エラー: " & Err.Description
                Exit Function
            End If
            On Error GoTo 0
            
            ' 条件が真なら対応する値を返す
            If bln_条件 Then
                IFSラッパー = 条件と値(i + 1)
                Exit Function
            End If
        Next i
        
        ' 全ての条件が偽で、デフォルト値も指定されていない場合
        IFSラッパー = CVErr(xlErrNA) ' #N/A エラー
    End If
End Function

'''---------------------------------------------------------
' 4. 日本語/英語など多言語対応の実装
'''---------------------------------------------------------
' ***************************************************************************************************
' モジュール: 多言語対応
' 機能: 異なる言語環境で動作するマクロのための多言語対応
' ***************************************************************************************************
Option Explicit

' 言語定数
Private Const INT_日本語 As Long = 1041
Private Const INT_英語 As Long = 1033
Private Const INT_中国語 As Long = 2052
Private Const INT_ドイツ語 As Long = 1031
Private Const INT_フランス語 As Long = 1036

' 言語リソーステーブル（実際のアプリでは外部ファイルから読み込むことも）
Private dict_メッセージ As Object ' Dictionary

'''---------------------------------------------------------
' 多言語対応のメイン処理
'''---------------------------------------------------------
Sub 多言語対応実装()
    ' 変数宣言
    Dim ws_情報 As Worksheet
    Dim lng_言語ID As Long
    Dim str_言語名 As String
    Dim str_メッセージ As String
    Dim i As Long
    
    ' 言語リソースの初期化
    Call 言語リソース初期化
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("多言語対応情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "多言語対応情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo 0
    
    ' 現在の言語設定を取得
    lng_言語ID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    ' 言語名の取得
    str_言語名 = 言語名取得(lng_言語ID)
    
    ' タイトル設定（リソースから取得）
    str_メッセージ = メッセージ取得("TITLE_多言語情報")
    ws_情報.Range("A1").Value = str_メッセージ
    ws_情報.Range("A1").Font.Size = 14
    ws_情報.Range("A1").Font.Bold = True
    
    ' 言語情報の表示
    str_メッセージ = メッセージ取得("LABEL_現在の言語")
    ws_情報.Range("A3").Value = str_メッセージ & ":"
    ws_情報.Range("B3").Value = str_言語名 & " (ID: " & lng_言語ID & ")"
    
    str_メッセージ = メッセージ取得("LABEL_Excelバージョン")
    ws_情報.Range("A4").Value = str_メッセージ & ":"
    ws_情報.Range("B4").Value = Application.Version
    
    str_メッセージ = メッセージ取得("LABEL_日付形式")
    ws_情報.Range("A5").Value = str_メッセージ & ":"
    ws_情報.Range("B5").Value = Application.International(xlDateOrder)
    
    str_メッセージ = メッセージ取得("LABEL_小数点記号")
    ws_情報.Range("A6").Value = str_メッセージ & ":"
    ws_情報.Range("B6").Value = Application.International(xlDecimalSeparator)
    
    str_メッセージ = メッセージ取得("LABEL_桁区切り記号")
    ws_情報.Range("A7").Value = str_メッセージ & ":"
    ws_情報.Range("B7").Value = Application.International(xlThousandsSeparator)
    
    ' 言語別メッセージサンプル
    str_メッセージ = メッセージ取得("TITLE_言語別メッセージ")
    ws_情報.Range("A9").Value = str_メッセージ & ":"
    ws_情報.Range("A9").Font.Bold = True
    
    ' サンプルメッセージを表示
    For i = 1 To 3
        str_メッセージ = メッセージ取得("SAMPLE_MESSAGE_" & i)
        ws_情報.Range("A" & (9 + i)).Value = str_メッセージ
    Next i
    
    ' 多言語対応のベストプラクティス
    str_メッセージ = メッセージ取得("TITLE_ベストプラクティス")
    ws_情報.Range("A14").Value = str_メッセージ & ":"
    ws_情報.Range("A14").Font.Bold = True
    
    ' ベストプラクティスの表示
    Dim arr_ベストプラクティス As Variant
    arr_ベストプラクティス = Array( _
        "BP_文字列リソース", _
        "BP_日付形式", _
        "BP_数値形式", _
        "BP_ファイル名", _
        "BP_シート名", _
        "BP_UIサイズ", _
        "BP_文字コード" _
    )
    
    For i = 0 To UBound(arr_ベストプラクティス)
        str_メッセージ = メッセージ取得(arr_ベストプラクティス(i))
        ws_情報.Range("A" & (15 + i)).Value = (i + 1) & ". " & str_メッセージ
    Next i
    
    ' 言語設定言語切り替えボタンの設置
    Dim btn_日本語 As Button
    Dim btn_英語 As Button
    
    ' 既存のボタンを削除
    On Error Resume Next
    For Each btn In ws_情報.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
    
    ' 日本語ボタン
    Set btn_日本語 = ws_情報.Buttons.Add(200, 200, 80, 30)
    With btn_日本語
        .Caption = "日本語"
        .Name = "btnJapanese"
        .OnAction = "言語切替_日本語"
    End With
    
    ' 英語ボタン
    Set btn_英語 = ws_情報.Buttons.Add(300, 200, 80, 30)
    With btn_英語
        .Caption = "English"
        .Name = "btnEnglish"
        .OnAction = "言語切替_英語"
    End With
    
    ' 位置調整
    btn_日本語.Top = ws_情報.Range("A24").Top
    btn_日本語.Left = ws_情報.Range("B24").Left
    
    btn_英語.Top = ws_情報.Range("A24").Top
    btn_英語.Left = ws_情報.Range("C24").Left
    
    ' 書式設定
    ws_情報.Columns("A:B").AutoFit
    
    ' アクティブ化
    ws_情報.Activate
    ws_情報.Range("A1").Select
End Sub

'''---------------------------------------------------------
' 言語リソースの初期化
'''---------------------------------------------------------
Private Sub 言語リソース初期化()
    ' Dictionaryオブジェクトの作成
    Set dict_メッセージ = CreateObject("Scripting.Dictionary")
    
    ' 日本語のメッセージ
    dict_メッセージ.Add INT_日本語 & "_TITLE_多言語情報", "多言語対応情報"
    dict_メッセージ.Add INT_日本語 & "_LABEL_現在の言語", "現在の言語"
    dict_メッセージ.Add INT_日本語 & "_LABEL_Excelバージョン", "Excelバージョン"
    dict_メッセージ.Add INT_日本語 & "_LABEL_日付形式", "日付形式"
    dict_メッセージ.Add INT_日本語 & "_LABEL_小数点記号", "小数点記号"
    dict_メッセージ.Add INT_日本語 & "_LABEL_桁区切り記号", "桁区切り記号"
    dict_メッセージ.Add INT_日本語 & "_TITLE_言語別メッセージ", "言語別メッセージサンプル"
    dict_メッセージ.Add INT_日本語 & "_SAMPLE_MESSAGE_1", "こんにちは、これは日本語のメッセージです。"
    dict_メッセージ.Add INT_日本語 & "_SAMPLE_MESSAGE_2", "日付: " & Format(Date, "yyyy年mm月dd日")
    dict_メッセージ.Add INT_日本語 & "_SAMPLE_MESSAGE_3", "金額: " & Format(12345.67, "#,##0.00") & " 円"
    dict_メッセージ.Add INT_日本語 & "_TITLE_ベストプラクティス", "多言語対応のベストプラクティス"
    dict_メッセージ.Add INT_日本語 & "_BP_文字列リソース", "すべての表示文字列をリソースとして外部化"
    dict_メッセージ.Add INT_日本語 & "_BP_日付形式", "地域によって異なる日付形式に注意（yyyy/mm/dd vs mm/dd/yyyy）"
    dict_メッセージ.Add INT_日本語 & "_BP_数値形式", "小数点と桁区切り記号の違いに注意（1,234.56 vs 1.234,56）"
    dict_メッセージ.Add INT_日本語 & "_BP_ファイル名", "ファイル名やパスに多言語文字を使わない"
    dict_メッセージ.Add INT_日本語 & "_BP_シート名", "シート名の言語依存性に注意（コードからは英語名で参照）"
    dict_メッセージ.Add INT_日本語 & "_BP_UIサイズ", "UIの翻訳で文字数が増えることに注意（余裕をもったサイズ設計）"
    dict_メッセージ.Add INT_日本語 & "_BP_文字コード", "ファイル入出力時の文字コードに注意（UTF-8推奨）"
    
    ' 英語のメッセージ
    dict_メッセージ.Add INT_英語 & "_TITLE_多言語情報", "Multilingual Support Information"
    dict_メッセージ.Add INT_英語 & "_LABEL_現在の言語", "Current Language"
    dict_メッセージ.Add INT_英語 & "_LABEL_Excelバージョン", "Excel Version"
    dict_メッセージ.Add INT_英語 & "_LABEL_日付形式", "Date Format"
    dict_メッセージ.Add INT_英語 & "_LABEL_小数点記号", "Decimal Separator"
    dict_メッセージ.Add INT_英語 & "_LABEL_桁区切り記号", "Thousands Separator"
    dict_メッセージ.Add INT_英語 & "_TITLE_言語別メッセージ", "Language-specific Messages"
    dict_メッセージ.Add INT_英語 & "_SAMPLE_MESSAGE_1", "Hello, this is an English message."
    dict_メッセージ.Add INT_英語 & "_SAMPLE_MESSAGE_2", "Date: " & Format(Date, "mm/dd/yyyy")
    dict_メッセージ.Add INT_英語 & "_SAMPLE_MESSAGE_3", "Amount: " & Format(12345.67, "$#,##0.00")
    dict_メッセージ.Add INT_英語 & "_TITLE_ベストプラクティス", "Multilingual Support Best Practices"
    dict_メッセージ.Add INT_英語 & "_BP_文字列リソース", "Externalize all display strings as resources"
    dict_メッセージ.Add INT_英語 & "_BP_日付形式", "Be aware of different date formats by region (yyyy/mm/dd vs mm/dd/yyyy)"
    dict_メッセージ.Add INT_英語 & "_BP_数値形式", "Pay attention to decimal and thousand separators (1,234.56 vs 1.234,56)"
    dict_メッセージ.Add INT_英語 & "_BP_ファイル名", "Avoid using multilingual characters in file names or paths"
    dict_メッセージ.Add INT_英語 & "_BP_シート名", "Be careful with language-dependent sheet names (reference by English name in code)"
    dict_メッセージ.Add INT_英語 & "_BP_UIサイズ", "Allow extra space for translations that may be longer than the original text"
    dict_メッセージ.Add INT_英語 & "_BP_文字コード", "Pay attention to character encodings when reading/writing files (UTF-8 recommended)"
    
    ' 他の言語のメッセージ（実際のアプリケーションではさらに追加）
End Sub

'''---------------------------------------------------------
' メッセージを取得する関数
' 引数:
'   Str_キー - メッセージのキー
' 戻り値:
'   String - 現在の言語に対応するメッセージ
'''---------------------------------------------------------
Private Function メッセージ取得(ByVal Str_キー As String) As String
    ' 変数宣言
    Dim lng_言語ID As Long
    Dim str_検索キー As String
    
    ' 現在の言語設定を取得
    lng_言語ID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    ' 言語IDとキーを組み合わせた検索キーを作成
    str_検索キー = lng_言語ID & "_" & Str_キー
    
    ' 言語リソースからメッセージを取得
    If dict_メッセージ.Exists(str_検索キー) Then
        メッセージ取得 = dict_メッセージ(str_検索キー)
    ElseIf dict_メッセージ.Exists(INT_英語 & "_" & Str_キー) Then
        ' 現在の言語のリソースがない場合は英語をフォールバックとして使用
        メッセージ取得 = dict_メッセージ(INT_英語 & "_" & Str_キー)
    Else
        ' キーが見つからない場合はキーをそのまま返す
        メッセージ取得 = Str_キー
    End If
End Function

'''---------------------------------------------------------
' 言語IDから言語名を取得する関数
'''---------------------------------------------------------
Private Function 言語名取得(ByVal Lng_言語ID As Long) As String
    ' 言語IDから言語名を判定
    Select Case Lng_言語ID
        Case INT_日本語: 言語名取得 = "日本語"
        Case INT_英語: 言語名取得 = "英語 (English)"
        Case INT_中国語: 言語名取得 = "中国語 (中文)"
        Case INT_ドイツ語: 言語名取得 = "ドイツ語 (Deutsch)"
        Case INT_フランス語: 言語名取得 = "フランス語 (Français)"
        Case Else: 言語名取得 = "その他の言語 (ID: " & Lng_言語ID & ")"
    End Select
End Function

'''---------------------------------------------------------
' 言語を日本語に切り替える（ボタン用）
'''---------------------------------------------------------
Public Sub 言語切替_日本語()
    ' ユーザーに対して言語切替の説明
    MsgBox "言語設定を変更するには、以下の手順で操作してください:" & vbCrLf & vbCrLf & _
           "1. [ファイル] メニュー → [オプション] を選択" & vbCrLf & _
           "2. ['''---------------------------------------------------------
' 1. 複数ユーザーによる同時アクセス制御
'''---------------------------------------------------------
Option Explicit

' 共有設定に関する定数
Private Const STR_ロックファイル拡張子 As String = ".lock"
Private Const STR_ユーザー情報ファイル As String = "user_access.log"
Private Const INT_ロックチェック間隔_ミリ秒 As Integer = 500
Private Const INT_最大待機時間_秒 As Integer = 30

'''---------------------------------------------------------
' ブックの共有ロック制御メイン処理
'''---------------------------------------------------------
Sub 共有ロック制御()
    ' 変数宣言
    Dim str_ロックファイル As String
    Dim bln_ロック取得 As Boolean
    Dim lng_開始時間 As Long
    
    ' 処理開始メッセージ
    Application.StatusBar = "共有環境を確認中..."
    
    ' ロックファイルパスの設定
    str_ロックファイル = ThisWorkbook.Path & "\" & ThisWorkbook.Name & STR_ロックファイル拡張子
    
    ' ロック取得を試行
    lng_開始時間 = Timer
    bln_ロック取得 = ロック取得(str_ロックファイル)
    
    ' ロック取得の待機ループ
    Do While Not bln_ロック取得
        ' 進捗表示を更新
        Application.StatusBar = "別のユーザーがファイルを使用中です。待機中... " & _
                              Format(Timer - lng_開始時間, "0") & "秒"
        
        ' 待機
        Sleep INT_ロックチェック間隔_ミリ秒
        DoEvents
        
        ' タイムアウトチェック
        If Timer - lng_開始時間 > INT_最大待機時間_秒 Then
            Application.StatusBar = False
            MsgBox "ファイルへのアクセスがタイムアウトしました。" & vbCrLf & _
                   "後ほど再試行するか、管理者に連絡してください。", vbExclamation
            Exit Sub
        End If
        
        ' 再度ロック取得を試行
        bln_ロック取得 = ロック取得(str_ロックファイル)
    Loop
    
    ' ロック取得成功
    Application.StatusBar = "ファイルへのアクセス権を取得しました。処理を開始します..."
    
    ' ユーザーアクセス記録
    Call ユーザーアクセス記録(ThisWorkbook.Name, "開始")
    
    ' ここに実際の処理を記述...
    Dim ws_情報 As Worksheet
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("共有アクセス情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "共有アクセス情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo 0
    
    ' 共有アクセス情報を表示
    アクセス情報表示 ws_情報
    
    ' 処理が終了したらロックを解放
    ロック解放 str_ロックファイル
    
    ' ユーザーアクセス記録
    Call ユーザーアクセス記録(ThisWorkbook.Name, "終了")
    
    Application.StatusBar = False
    MsgBox "処理が完了しました。", vbInformation
End Sub

'''---------------------------------------------------------
' ロックファイルを作成してロックを取得する
' 引数:
'   Str_ロックファイル - ロックファイルのパス
' 戻り値:
'   Boolean - ロック取得結果（True=成功、False=失敗）
'''---------------------------------------------------------
Private Function ロック取得(ByVal Str_ロックファイル As String) As Boolean
    ' 変数宣言
    Dim int_ファイル番号 As Integer
    Dim obj_FSO As Object
    
    ' 初期化
    ロック取得 = False
    
    ' ロックファイルが既に存在するか確認
    On Error Resume Next
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' ロックファイルが存在する場合はロック取得失敗
    If obj_FSO.FileExists(Str_ロックファイル) Then
        ' ロックファイルの最終更新時刻を確認
        Dim obj_ファイル As Object
        Set obj_ファイル = obj_FSO.GetFile(Str_ロックファイル)
        
        ' ロックファイルが古い場合（10分以上経過）は強制的に削除
        If DateDiff("n", obj_ファイル.DateLastModified, Now) > 10 Then
            obj_FSO.DeleteFile Str_ロックファイル, True
        Else
            ' 有効なロックが存在するので失敗
            Exit Function
        End If
    End If
    
    ' ロックファイルを作成
    int_ファイル番号 = FreeFile
    Open Str_ロックファイル For Output As #int_ファイル番号
    Print #int_ファイル番号, "ロック取得者: " & Application.UserName
    Print #int_ファイル番号, "ロック取得時刻: " & Now
    Print #int_ファイル番号, "コンピュータ名: " & Environ("COMPUTERNAME")
    Close #int_ファイル番号
    
    ' ロックファイルが正常に作成できたらロック取得成功
    If Err.Number = 0 Then
        ロック取得 = True
    End If
    On Error GoTo 0
End Function

'''---------------------------------------------------------
' ロックファイルを削除してロックを解放する
' 引数:
'   Str_ロックファイル - ロックファイルのパス
'''---------------------------------------------------------
Private Sub ロック解放(ByVal Str_ロックファイル As String)
    ' ロックファイルが存在する場合は削除
    On Error Resume Next
    Dim obj_FSO As Object
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    If obj_FSO.FileExists(Str_ロックファイル) Then
        obj_FSO.DeleteFile Str_ロックファイル, True
    End If
    On Error GoTo 0
End Sub

'''---------------------------------------------------------
' ユーザーアクセスを記録する
' 引数:
'   Str_ファイル名 - アクセスしたファイル名
'   Str_操作 - 実行した操作（開始/終了など）
'''---------------------------------------------------------
Private Sub ユーザーアクセス記録(ByVal Str_ファイル名 As String, ByVal Str_操作 As String)
    ' 変数宣言
    Dim str_ログファイル As String
    Dim int_ファイル番号 As Integer
    
    ' ログファイルパスの設定
    str_ログファイル = ThisWorkbook.Path & "\" & STR_ユーザー情報ファイル
    
    ' ログファイルに追記
    On Error Resume Next
    int_ファイル番号 = FreeFile
    Open str_ログファイル For Append As #int_ファイル番号
    Print #int_ファイル番号, Format(Now, "yyyy/mm/dd hh:nn:ss") & "," & _
                           Application.UserName & "," & _
                           Environ("COMPUTERNAME") & "," & _
                           Str_ファイル名 & "," & _
                           Str_操作
    Close #int_ファイル番号
    On Error GoTo 0
End Sub

'''---------------------------------------------------------
' アクセス情報をワークシートに表示する
' 引数:
'   Ws_情報 - 情報を表示するワークシート
'''---------------------------------------------------------
Private Sub アクセス情報表示(ByVal Ws_情報 As Worksheet)
    ' 変数宣言
    Dim str_ログファイル As String
    Dim int_ファイル番号 As Integer
    Dim str_行 As String
    Dim arr_ログデータ() As String
    Dim i As Long
    Dim lng_行数 As Long
    
    ' タイトル設定
    Ws_情報.Range("A1").Value = "共有ファイルアクセス記録"
    Ws_情報.Range("A1").Font.Size = 14
    Ws_情報.Range("A1").Font.Bold = True
    
    ' ヘッダー行の設定
    Ws_情報.Range("A3").Value = "日時"
    Ws_情報.Range("B3").Value = "ユーザー名"
    Ws_情報.Range("C3").Value = "コンピュータ名"
    Ws_情報.Range("D3").Value = "ファイル名"
    Ws_情報.Range("E3").Value = "操作"
    Ws_情報.Range("A3:E3").Font.Bold = True
    
    ' ログファイルパスの設定
    str_ログファイル = ThisWorkbook.Path & "\" & STR_ユーザー情報ファイル
    
    ' ログファイルが存在する場合は読み込み
    If Dir(str_ログファイル) <> "" Then
        ' ログファイルを開く
        int_ファイル番号 = FreeFile
        Open str_ログファイル For Input As #int_ファイル番号
        
        ' 行数のカウント
        lng_行数 = 0
        Do Until EOF(int_ファイル番号)
            Line Input #int_ファイル番号, str_行
            lng_行数 = lng_行数 + 1
        Loop
        Close #int_ファイル番号
        
        ' 配列の初期化
        ReDim arr_ログデータ(1 To lng_行数, 1 To 5)
        
        ' 再度ファイルを開いてデータを配列に格納
        int_ファイル番号 = FreeFile
        Open str_ログファイル For Input As #int_ファイル番号
        
        i = 1
        Do Until EOF(int_ファイル番号)
            Line Input #int_ファイル番号, str_行
            
            ' カンマ区切りの行をフィールドに分割
            Dim arr_フィールド As Variant
            arr_フィールド = Split(str_行, ",")
            
            ' 配列にデータを格納（フィールド数が不足している場合のエラー回避）
            If UBound(arr_フィールド) >= 4 Then
                arr_ログデータ(i, 1) = arr_フィールド(0) ' 日時
                arr_ログデータ(i, 2) = arr_フィールド(1) ' ユーザー名
                arr_ログデータ(i, 3) = arr_フィールド(2) ' コンピュータ名
                arr_ログデータ(i, 4) = arr_フィールド(3) ' ファイル名
                arr_ログデータ(i, 5) = arr_フィールド(4) ' 操作
            End If
            
            i = i + 1
        Loop
        Close #int_ファイル番号
        
        ' 最新100件のみ表示
        Dim lng_開始インデックス As Long
        If lng_行数 > 100 Then
            lng_開始インデックス = lng_行数 - 99
        Else
            lng_開始インデックス = 1
        End If
        
        ' データをワークシートに表示
        For i = lng_開始インデックス To lng_行数
            Ws_情報.Cells(i - lng_開始インデックス + 4, 1).Value = arr_ログデータ(i, 1)
            Ws_情報.Cells(i - lng_開始インデックス + 4, 2).Value = arr_ログデータ(i, 2)
            Ws_情報.Cells(i - lng_開始インデックス + 4, 3).Value = arr_ログデータ(i, 3)
            Ws_情報.Cells(i - lng_開始インデックス + 4, 4).Value = arr_ログデータ(i, 4)
            Ws_情報.Cells(i - lng_開始インデックス + 4, 5).Value = arr_ログデータ(i, 5)
            
            ' 自分のアクセス行を強調表示
            If arr_ログデータ(i, 2) = Application.UserName Then
                Ws_情報.Range("A" & (i - lng_開始インデックス + 4) & ":E" & (i - lng_開始インデックス + 4)).Interior.Color = RGB(220, 240, 255) ' 薄い青色
            End If
        Next i
        
        ' 書式設定
        Ws_情報.Columns("A:E").AutoFit
        Ws_情報.Range("A3:E" & (lng_行数 - lng_開始インデックス + 4)).Borders.LineStyle = xlContinuous
    Else
        ' ログファイルが存在しない場合
        Ws_情報.Range("A4").Value = "ログファイルが見つかりません。"
    End If
    
    ' 共有環境のベストプラクティス
    Ws_情報.Range("A" & lng_行数 + 6).Value = "共有環境での操作ガイドライン:"
    Ws_情報.Range("A" & lng_行数 + 6).Font.Bold = True
    
    Ws_情報.Range("A" & lng_行数 + 7).Value = "1. 編集前に必ず最新版を取得（他のユーザーの変更を確認）"
    Ws_情報.Range("A" & lng_行数 + 8).Value = "2. 作業が完了したら速やかに保存して閉じる"
    Ws_情報.Range("A" & lng_行数 + 9).Value = "3. 長時間の編集を行う場合は事前に他のユーザーに通知"
    Ws_情報.Range("A" & lng_行数 + 10).Value = "4. 同じシートを複数人で同時に編集しない"
    Ws_情報.Range("A" & lng_行数 + 11).Value = "5. 問題が発生した場合は直ちにシステム管理者に連絡"
End Sub

' Sleep API関数（待機用）
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'''---------------------------------------------------------
' 2. 32bit/64bit環境の違いに対応する手法
'''---------------------------------------------------------
' ***************************************************************************************************
' モジュール: 32bit/64bit互換性対応
' 機能: 異なるビット環境で正しく動作するためのコード手法
' ***************************************************************************************************
Option Explicit

' API宣言（32bit/64bit両対応）
#If VBA7 Then
    ' Excel 2010以降の64bit対応バージョン
    Private Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" ( _
        ByVal lpBuffer As String, ByVal nSize As Long) As Long
    
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, ByRef nSize As Long) As Long
        
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long
#Else
    ' Excel 2007以前の32bit専用バージョン
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" ( _
        ByVal lpBuffer As String, ByVal nSize As Long) As Long
    
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
        ByVal lpBuffer As String, ByRef nSize As Long) As Long
        
    Private Declare Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long
#End If

' GetSystemMetrics用の定数
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

'''---------------------------------------------------------
' 32bit/64bit環境の違いに対応するコードサンプル
'''---------------------------------------------------------
Sub ビット環境対応実装()
    ' 変数宣言
    Dim ws_情報 As Worksheet
    Dim str_Windowsディレクトリ As String
    Dim str_ユーザー名 As String
    Dim lng_画面幅 As Long
    Dim lng_画面高さ As Long
    Dim bln_64ビットExcel As Boolean
    Dim bln_64ビットOS As Boolean
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("ビット環境情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "ビット環境情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo 0
    
    ' タイトル設定
    ws_情報.Range("A1").Value = "32bit/64bit環境情報"
    ws_情報.Range("A1").Font.Size = 14
    ws_情報.Range("A1").Font.Bold = True
    
    ' 実行環境のビット数チェック
    #If Win64 Then
        bln_64ビットExcel = True
    #Else
        bln_64ビットExcel = False
    #End If
    
    ' OSのビット数チェック（間接的な方法）
    Dim str_PROCESSOR_ARCHITECTURE As String
    str_PROCESSOR_ARCHITECTURE = Environ("PROCESSOR_ARCHITECTURE")
    bln_64ビットOS = (str_PROCESSOR_ARCHITECTURE = "AMD64" Or str_PROCESSOR_ARCHITECTURE = "IA64")
    
    ' 環境情報の表示
    ws_情報.Range("A3").Value = "Excel実行環境:"
    ws_情報.Range("B3").Value = IIf(bln_64ビットExcel, "64ビット", "32ビット")
    
    ws_情報.Range("A4").Value = "OSアーキテクチャ:"
    ws_情報.Range("B4").Value = IIf(bln_64ビットOS, "64ビット", "32ビット")
    
    ws_情報.Range("A5").Value = "Excelバージョン:"
    ws_情報.Range("B5").Value = Application.Version
    
    ' Windowsディレクトリの取得
    str_Windowsディレクトリ = 文字列API呼び出し_GetWindowsDirectory()
    ws_情報.Range("A6").Value = "Windowsディレクトリ:"
    ws_情報.Range("B6").Value = str_Windowsディレクトリ
    
    ' ユーザー名の取得
    str_ユーザー名 = 文字列API呼び出し_GetUserName()
    ws_情報.Range("A7").Value = "ユーザー名 (API):"
    ws_情報.Range("B7").Value = str_ユーザー名
    
    ' 画面解像度の取得
    lng_画面幅 = GetSystemMetrics(SM_CXSCREEN)
    lng_画面高さ = GetSystemMetrics(SM_CYSCREEN)
    ws_情報.Range("A8").Value = "画面解像度:"
    ws_情報.Range("B8").Value = lng_画面幅 & " x " & lng_画面高さ
    
    ' 注意事項の表示
    Dim i As Long
    Dim arr_注意事項 As Variant
    
    arr_注意事項 = Array( _
        "32bit/64bit互換性のためのコードガイドライン:", _
        "", _
        "1. API宣言には必ず条件付きコンパイルディレクティブ (#If VBA7 Then) を使用", _
        "2. 64bit環境ではポインタを扱う宣言に PtrSafe キーワードを追加", _
        "3. 32bit版ではLong型、64bit版ではLongLong型を使用する場合は条件分岐", _
        "4. Windows APIのハンドル（HANDLE, HWND等）を扱う場合は特に注意", _
        "5. バッファサイズの計算が複雑になる場合があるので余裕を持って確保", _
        "6. メモリ操作系のAPIは特に注意（ポインタ演算や配列操作）", _
        "7. 古いCOMコンポーネントは64bit環境で動作しない場合がある" _
    )
    
    ws_情報.Range("A10").Value = arr_注意事項(0)
    ws_情報.Range("A10").Font.Bold = True
    
    For i = 1 To UBound(arr_注意事項)
        ws_情報.Range("A" & (10 + i)).Value = arr_注意事項(i)
    Next i
    
    ' 32bit/64bit対応APIサンプルコード
    ws_情報.Range("A20").Value = "32bit/64bit互換API宣言の例:"
    ws_情報.Range("A20").Font.Bold = True
    
    Dim arr_コード例 As Variant
    arr_コード例 = Array( _
        "#If VBA7 Then", _
        "    Private Declare PtrSafe Function GetWindowsDirectory Lib ""kernel32"" Alias ""GetWindowsDirectoryA"" (", _
        "        ByVal lpBuffer As String, ByVal nSize As Long) As Long", _
        "#Else", _
        "    Private Declare Function GetWindowsDirectory Lib ""kernel32"" Alias ""GetWindowsDirectoryA"" (", _
        "        ByVal lpBuffer As String, ByVal nSize As Long) As Long", _
        "#End If", _
        "", _
        "' 64bitと32bit両対応のポインタ型", _
        "#If VBA7 Then", _
        "    Dim ptr As LongPtr  ' 64bitでは64ビット幅、32bitでは32ビット幅", _
        "#Else", _
        "    Dim ptr As Long     ' 32bitのみ", _
        "#End If" _
    )
    
    For i = 0 To UBound(arr_コード例)
        ws_情報.Range("A" & (21 + i)).Value = arr_コード例(i)
    Next i
    
    ' 書式設定
    ws_情報.Columns("A:B").AutoFit
    
    ' アクティブ化
    ws_情報.Activate
    ws_情報.Range("A1").Select
End Sub

'''---------------------------------------------------------
' Windowsディレクトリを取得する関数（32bit/64bit両対応）
'''---------------------------------------------------------
Private Function 文字列API呼び出し_GetWindowsDirectory() As String
    ' 変数宣言
    Dim str_バッファ As String
    Dim lng_バッファサイズ As Long
    Dim lng_戻り値 As Long
    
    ' バッファの初期化（十分なサイズを確保）
    lng_バッファサイズ = 260 ' MAX_PATH
    str_バッファ = String(lng_バッファサイズ, 0)
    
    ' API呼び出し
    lng_戻り値 = GetWindowsDirectory(str_バッファ, lng_バッファサイズ)
    
    ' 戻り値が0より大きい場合は成功
    If lng_戻り値 > 0 Then
        ' 結果を取得（null終端文字を除去）
        文字列API呼び出し_GetWindowsDirectory = Left(str_バッファ, lng_戻り値)
    Else
        ' エラー時は空文字を返す
        文字列API呼び出し_GetWindowsDirectory = ""
    End If
End Function

'''---------------------------------------------------------
' ユーザー名を取得する関数（32bit/64bit両対応）
'''---------------------------------------------------------
Private Function 文字列API呼び出し_GetUserName() As String
    ' 変数宣言
    Dim str_バッファ As String
    Dim lng_バッファサイズ As Long
    Dim lng_戻り値 As Long
    
    ' バッファの初期化
    lng_バッファサイズ = 256
    str_バッファ = String(lng_バッファサイズ, 0)
    
    ' API呼び出し
    lng_戻り値 = GetUserName(str_バッファ, lng_バッファサイズ)
    
    ' 戻り値が0でない場合は成功
    If lng_戻り値 <> 0 Then
        ' 結果を取得（null終端文字を除去）
        文字列API呼び出し_GetUserName = Left(str_バッファ, InStr(str_バッファ, Chr(0)) - 1)
    Else
        ' エラー時は空文字を返す
        文字列API呼び出し_GetUserName = ""
    End If
End Function

'''---------------------------------------------------------
' 3. Excel各バージョン間の互換性対応
'''---------------------------------------------------------
' ***************************************************************************************************
' モジュール: Excel互換性対応
' 機能: 異なるExcelバージョン間で互換性を保つためのコード設計
' ***************************************************************************************************
Option Explicit

' バージョン依存の挙動の違いに関する定数
Private Const STR_最小対応バージョン As String = "14.0" ' Excel 2010
Private Const STR_最新対応バージョン As String = "16.0" ' Excel 2016/365

'''---------------------------------------------------------
' Excel互換性対応のメイン処理
'''---------------------------------------------------------
Sub Excel互換性対応()
    ' 変数宣言
    Dim ws_情報 As Worksheet
    Dim dbl_Excelバージョン As Double
    Dim str_バージョン名 As String
    Dim bln_対応バージョン As Boolean
    Dim i As Long
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("Excel互換性情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "Excel互換性情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo 0
    
    ' タイトル設定
    ws_情報.Range("A1").Value = "Excel互換性情報"
    ws_情報.Range("A1").Font.Size = 14
    ws_情報.Range("A1").Font.Bold = True
    
    ' Excelバージョン情報
    dbl_Excelバージョン = CDbl(Application.Version)
    
    ' バージョン名の取得
    str_バージョン名 = Excelバージョン名取得(dbl_Excelバージョン)
    
    ' 対応バージョンかどうかチェック
    bln_対応バージョン = (dbl_Excelバージョン >= CDbl(STR_最小対応バージョン))
    
    ' バージョン情報の表示
    ws_情報.Range("A3").Value = "Excelバージョン:"
    ws_情報.Range("B3").Value = Application.Version & " (" & str_バージョン名 & ")"
    
    ws_情報.Range("A4").Value = "ビルド番号:"
    ws_情報.Range("B4").Value = Application.Build
    
    ws_情報.Range("A5").Value = "対応状況:"
    
    ' 対応状況の表示
    If bln_対応バージョン Then
        ws_情報.Range("B5").Value = "対応バージョンです"
        ws_情報.Range("B5").Font.Color = RGB(0, 128, 0) ' 緑色
    Else
        ws_情報.Range("B5").Value = "非対応バージョンです（" & STR_最小対応バージョン & "以上が必要）"
        ws_情報.Range("B5").Font.Color = RGB(192, 0, 0) ' 赤色
    End If
    
    ' 機能対応表の作成
    ws_情報.Range("A7").Value = "Excel各バージョンの機能対応表:"
    ws_情報.Range("A7").Font.Bold = True
    
    ' ヘッダー行
    ws_情報.Range("A8").Value = "機能名"
    ws_情報.Range("B8").Value = "Excel 2007"
    ws_情報.Range("C8").Value = "Excel 2010"
    ws_情報.Range("D8").Value = "Excel 2013"
    ws_情報.Range("E8").Value = "Excel 2016/365"
    
    ' ヘッダー書式設定
    ws_情報.Range("A8:E8").Font.Bold = True
    ws_情報.Range("A8:E8").Interior.Color = RGB(220, 220, 220)
    
    ' 機能対応データ
    Dim arr_機能 As Variant
    arr_機能 = Array( _
        Array("スライサー", "×", "○", "○", "○"), _
        Array("パワーピボット", "×", "△", "○", "○"), _
        Array("フラッシュフィル", "×", "×", "○", "○"), _
        Array("IFS関数", "×", "×", "×", "○"), _
        Array("TEXTJOIN関数", "×", "×", "×", "○"), _
        Array("動的配列関数", "×", "×", "×", "△"), _
        Array("ストックデータ型", "×", "×", "×", "△") _
    )
    
    ' 機能対応表の作成
    For i = 0 To UBound(arr_機能)
        ws_情報.Cells(9 + i, 1).Value = arr_機能(i)(0)
        
        For j = 1 To 4
            ws_情報.Cells(9 + i, j + 1).Value = arr_機能(i)(j)
            ws_情報.Cells(9 + i, j + 1).HorizontalAlignment = xlCenter
        Next j
    Next i
End Sub
