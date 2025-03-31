
'''---------------------------------------------------------
' 1. Dictionary活用による検索・集計の高速化
'''---------------------------------------------------------
Sub Dictionary活用サンプル()
    ' 参照設定: Microsoft Scripting Runtime
    ' VBE→ツール→参照設定から追加する必要あり
    
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim dict_集計 As Object ' または Dictionary として宣言可能（参照設定必要）
    Dim lng_最終行 As Long
    Dim lng_行 As Long
    Dim str_キー As String
    Dim dbl_値 As Double
    Dim var_キー As Variant
    Dim lng_出力行 As Long
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("データ")
    
    ' Dictionaryオブジェクトの作成（エラー対策のため Late Binding を使用）
    Set dict_集計 = CreateObject("Scripting.Dictionary")
    dict_集計.CompareMode = vbTextCompare ' 大文字小文字を区別しない
    
    ' 最終行を取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    
    ' データをDictionaryに格納（例：A列を顧客ID、B列を売上とする）
    For lng_行 = 2 To lng_最終行 ' ヘッダー行をスキップ
        str_キー = CStr(ws_データ.Cells(lng_行, 1).Value) ' 顧客ID
        dbl_値 = CDbl(ws_データ.Cells(lng_行, 2).Value)   ' 売上額
        
        ' 既にキーが存在するかチェック
        If dict_集計.Exists(str_キー) Then
            ' 既存の値に加算
            dict_集計(str_キー) = dict_集計(str_キー) + dbl_値
        Else
            ' 新しいキーを追加
            dict_集計.Add str_キー, dbl_値
        End If
    Next lng_行
    
    ' 結果を別シートに出力
    Dim ws_結果 As Worksheet
    On Error Resume Next
    Set ws_結果 = ThisWorkbook.Worksheets("集計結果")
    If ws_結果 Is Nothing Then
        Set ws_結果 = ThisWorkbook.Worksheets.Add
        ws_結果.Name = "集計結果"
    End If
    On Error GoTo 0
    
    ' ヘッダー行の設定
    ws_結果.Cells(1, 1).Value = "顧客ID"
    ws_結果.Cells(1, 2).Value = "合計売上"
    
    ' Dictionaryの内容を出力
    lng_出力行 = 2
    For Each var_キー In dict_集計.Keys
        ws_結果.Cells(lng_出力行, 1).Value = var_キー
        ws_結果.Cells(lng_出力行, 2).Value = dict_集計(var_キー)
        lng_出力行 = lng_出力行 + 1
    Next var_キー
    
    ' 書式設定
    ws_結果.Range("A1:B1").Font.Bold = True
    ws_結果.Range("B2:B" & lng_出力行 - 1).NumberFormat = "#,##0"
    ws_結果.Columns("A:B").AutoFit
    
    ' 後処理
    Set dict_集計 = Nothing
    
    MsgBox "集計が完了しました。" & dict_集計.Count & "件の顧客データを処理しました。", vbInformation
End Sub

'''---------------------------------------------------------
' 2. 名前付き範囲の動的操作テクニック
'''---------------------------------------------------------
Sub 名前付き範囲操作サンプル()
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim rng_データ範囲 As Range
    Dim rng_新範囲 As Range
    Dim lng_最終行 As Long
    Dim str_名前 As String
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("データ")
    
    ' 最終行を取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    
    ' 動的範囲の設定（A1からC列の最終行まで）
    Set rng_データ範囲 = ws_データ.Range("A1:C" & lng_最終行)
    
    ' 1. 既存の名前付き範囲を更新する
    str_名前 = "データ範囲"
    
    On Error Resume Next
    ' 名前付き範囲が存在するか確認し、存在すれば削除
    If Not ThisWorkbook.Names(str_名前) Is Nothing Then
        ThisWorkbook.Names(str_名前).Delete
    End If
    On Error GoTo 0
    
    ' 新しい名前付き範囲を作成
    ThisWorkbook.Names.Add Name:=str_名前, RefersTo:=rng_データ範囲
    
    ' 2. 動的な参照を使った名前付き範囲（OFFSET関数を使用）
    str_名前 = "動的データ範囲"
    
    On Error Resume Next
    If Not ThisWorkbook.Names(str_名前) Is Nothing Then
        ThisWorkbook.Names(str_名前).Delete
    End If
    On Error GoTo 0
    
    ' OFFSET関数を使って動的範囲を作成
    ' =OFFSET(データ!$A$1,0,0,COUNTA(データ!$A:$A),3)
    ThisWorkbook.Names.Add Name:=str_名前, _
        RefersTo:="=OFFSET(" & ws_データ.Name & "!$A$1,0,0,COUNTA(" & _
                 ws_データ.Name & "!$A:$A),3)"
    
    ' 3. テーブル列を名前付き範囲として定義
    ' テーブルが存在するか確認
    If ws_データ.ListObjects.Count > 0 Then
        ' 最初のテーブルを使用
        Dim tbl_データ As ListObject
        Set tbl_データ = ws_データ.ListObjects(1)
        
        ' テーブルの列を名前付き範囲として定義
        For Each col In tbl_データ.ListColumns
            str_名前 = "テーブル_" & col.Name
            
            On Error Resume Next
            If Not ThisWorkbook.Names(str_名前) Is Nothing Then
                ThisWorkbook.Names(str_名前).Delete
            End If
            On Error GoTo 0
            
            ThisWorkbook.Names.Add Name:=str_名前, RefersTo:=col.DataBodyRange
        Next col
    End If
    
    ' 4. 名前付き範囲を使ったデータ取得
    If Not ThisWorkbook.Names("データ範囲") Is Nothing Then
        ' 名前付き範囲を取得
        Set rng_新範囲 = ThisWorkbook.Names("データ範囲").RefersToRange
        
        ' 名前付き範囲を使った処理例
        Dim var_データ As Variant
        var_データ = rng_新範囲.Value
        
        MsgBox "データ範囲のサイズ: " & rng_新範囲.Rows.Count & "行 x " & _
               rng_新範囲.Columns.Count & "列", vbInformation
    End If
    
    ' 5. 名前付き範囲の一覧を表示
    Dim str_名前一覧 As String
    Dim nm As Name
    
    str_名前一覧 = "このブックに定義されている名前付き範囲:" & vbCrLf & vbCrLf
    
    For Each nm In ThisWorkbook.Names
        str_名前一覧 = str_名前一覧 & nm.Name & " - " & nm.RefersTo & vbCrLf
    Next nm
    
    MsgBox str_名前一覧, vbInformation, "名前付き範囲一覧"
End Sub

'''---------------------------------------------------------
' 3. テーブル（ListObject）の効率的な操作
'''---------------------------------------------------------
Sub テーブル操作サンプル()
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim tbl_データ As ListObject
    Dim rng_データ範囲 As Range
    Dim var_テーブルデータ As Variant
    Dim lng_行数 As Long
    Dim lng_列数 As Long
    Dim lng_行 As Long
    Dim lng_列 As Long
    Dim arr_ヘッダー() As String
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("データ")
    
    ' 既存のテーブルがあるか確認
    If ws_データ.ListObjects.Count > 0 Then
        ' 最初のテーブルを取得
        Set tbl_データ = ws_データ.ListObjects(1)
    Else
        ' データ範囲を取得（A1からデータがある最終セルまで）
        Set rng_データ範囲 = ws_データ.Range("A1").CurrentRegion
        
        ' データ範囲をテーブルに変換
        Set tbl_データ = ws_データ.ListObjects.Add(xlSrcRange, rng_データ範囲, , xlYes)
        tbl_データ.Name = "データテーブル"
    End If
    
    ' 1. テーブルの基本情報を取得
    MsgBox "テーブル名: " & tbl_データ.Name & vbCrLf & _
           "行数: " & tbl_データ.ListRows.Count & vbCrLf & _
           "列数: " & tbl_データ.ListColumns.Count, vbInformation, "テーブル情報"
    
    ' 2. テーブルデータを配列に取得
    var_テーブルデータ = tbl_データ.DataBodyRange.Value
    lng_行数 = UBound(var_テーブルデータ, 1)
    lng_列数 = UBound(var_テーブルデータ, 2)
    
    ' ヘッダー配列の初期化
    ReDim arr_ヘッダー(1 To lng_列数)
    For lng_列 = 1 To lng_列数
        arr_ヘッダー(lng_列) = tbl_データ.HeaderRowRange(1, lng_列).Value
    Next lng_列
    
    ' 3. テーブルに新しい行を追加
    Dim lst_新行 As ListRow
    Set lst_新行 = tbl_データ.ListRows.Add
    
    ' 新しい行にデータを設定
    For lng_列 = 1 To lng_列数
        lst_新行.Range(1, lng_列).Value = "新規データ_" & lng_列
    Next lng_列
    
    ' 4. テーブルフィルタリングとソート
    ' フィルターを有効化
    tbl_データ.ShowAutoFilter = True
    
    ' 特定の列でフィルタリング（例：1列目で「東京」を含むデータをフィルタリング）
    tbl_データ.Range.AutoFilter Field:=1, Criteria1:="*東京*", Operator:=xlFilterValues
    
    ' フィルタリングをクリア
    tbl_データ.Range.AutoFilter
    
    ' テーブルをソート（例：2列目を昇順でソート）
    With tbl_データ.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl_データ.ListColumns(2).Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    ' 5. テーブルスタイルの変更
    tbl_データ.TableStyle = "TableStyleMedium2"
    
    ' 6. 構造化参照を使った式の追加
    If tbl_データ.ListColumns.Count >= 3 Then
        ' 既に計算列が存在するか確認
        On Error Resume Next
        Set calc_col = tbl_データ.ListColumns("計算結果")
        On Error GoTo 0
        
        If calc_col Is Nothing Then
            ' 計算列の追加
            Set calc_col = tbl_データ.ListColumns.Add
            calc_col.Name = "計算結果"
            
            ' 構造化参照を使った計算式を設定
            ' 例：1列目と2列目の値を足し算
            calc_col.DataBodyRange.FormulaR1C1 = "=[@[" & arr_ヘッダー(1) & "]]+[@[" & arr_ヘッダー(2) & "]]"
        End If
    End If
    
    MsgBox "テーブル操作が完了しました。", vbInformation
End Sub

'''---------------------------------------------------------
' 4. 大規模データセットのソート・フィルタリング手法
'''---------------------------------------------------------
Sub 大規模データ処理サンプル()
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim rng_データ範囲 As Range
    Dim lng_最終行 As Long
    Dim lng_最終列 As Long
    Dim arr_データ() As Variant
    Dim lng_行 As Long
    Dim lng_列 As Long
    Dim dbl_開始時間 As Double
    
    ' 処理時間の計測開始
    dbl_開始時間 = Timer
    
    ' 画面更新を停止（高速化）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("大規模データ")
    
    ' データ範囲の取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    lng_最終列 = ws_データ.Cells(1, ws_データ.Columns.Count).End(xlToLeft).Column
    
    If lng_最終行 <= 1 Then
        MsgBox "データが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Set rng_データ範囲 = ws_データ.Range(ws_データ.Cells(1, 1), ws_データ.Cells(lng_最終行, lng_最終列))
    
    ' 1. メモリ上でデータを処理するために配列に格納
    arr_データ = rng_データ範囲.Value
    
    ' 2. メモリ上でのカスタムソート（例：2列目を基準にクイックソート）
    If lng_最終行 > 1 Then  ' ヘッダー以外のデータがある場合
        Call クイックソート配列(arr_データ, 2, 2, lng_最終行)
    End If
    
    ' 3. ソートしたデータをシートに書き戻す
    rng_データ範囲.Value = arr_データ
    
    ' 4. 高度なフィルタリング（例：条件に合うデータを別シートに抽出）
    Dim ws_フィルタ結果 As Worksheet
    
    ' フィルタ結果シートを作成または取得
    On Error Resume Next
    Set ws_フィルタ結果 = ThisWorkbook.Worksheets("フィルタ結果")
    On Error GoTo 0
    
    If ws_フィルタ結果 Is Nothing Then
        Set ws_フィルタ結果 = ThisWorkbook.Worksheets.Add(After:=ws_データ)
        ws_フィルタ結果.Name = "フィルタ結果"
    Else
        ws_フィルタ結果.Cells.Clear
    End If
    
    ' フィルタ条件シートを作成（一時的）
    Dim ws_条件 As Worksheet
    On Error Resume Next
    Set ws_条件 = ThisWorkbook.Worksheets("条件")
    If ws_条件 Is Nothing Then
        Set ws_条件 = ThisWorkbook.Worksheets.Add
        ws_条件.Name = "条件"
    End If
    ws_条件.Visible = xlSheetHidden
    On Error GoTo 0
    
    ' 条件を設定（例：3列目が100以上）
    ws_条件.Cells.Clear
    ws_条件.Range("A1").Value = arr_データ(1, 3)  ' ヘッダー名
    ws_条件.Range("A2").Value = ">100"       ' 条件
    
    ' 高度なフィルターを実行
    rng_データ範囲.AdvancedFilter Action:=xlFilterCopy, _
                     CriteriaRange:=ws_条件.Range("A1:A2"), _
                     CopyToRange:=ws_フィルタ結果.Range("A1"), _
                     Unique:=False
    
    ' 5. データの集計と分析
    Dim ws_集計 As Worksheet
    
    On Error Resume Next
    Set ws_集計 = ThisWorkbook.Worksheets("集計結果")
    If ws_集計 Is Nothing Then
        Set ws_集計 = ThisWorkbook.Worksheets.Add(After:=ws_フィルタ結果)
        ws_集計.Name = "集計結果"
    Else
        ws_集計.Cells.Clear
    End If
    On Error GoTo 0
    
    ' ピボットテーブルの作成
    Dim pvt_集計 As PivotTable
    Dim pvt_キャッシュ As PivotCache
    
    ' ピボットキャッシュの作成
    Set pvt_キャッシュ = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rng_データ範囲)
    
    ' ピボットテーブルの作成
    Set pvt_集計 = pvt_キャッシュ.CreatePivotTable( _
        TableDestination:=ws_集計.Range("A3"), _
        TableName:="データ集計")
    
    ' フィールドの追加（例：1列目を行、2列目を列、3列目を値）
    With pvt_集計
        .AddFields RowFields:=Array(arr_データ(1, 1)), _
                  ColumnFields:=Array(arr_データ(1, 2))
        
        ' 値フィールドの追加
        .AddDataField pvt_集計.PivotFields(arr_データ(1, 3)), _
                      "合計 " & arr_データ(1, 3), xlSum
    End With
    
    ' 設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 結果の表示
    MsgBox "処理が完了しました。" & vbCrLf & _
           "処理時間: " & Format(Timer - dbl_開始時間, "0.00") & " 秒" & vbCrLf & _
           "処理行数: " & lng_最終行 - 1, vbInformation
End Sub

' クイックソートアルゴリズム（2次元配列用）
Private Sub クイックソート配列(arr_データ() As Variant, lng_ソート列 As Long, lng_開始 As Long, lng_終了 As Long)
    Dim lng_左 As Long
    Dim lng_右 As Long
    Dim var_ピボット As Variant
    Dim var_一時 As Variant
    Dim lng_列 As Long
    
    ' 基底条件
    If lng_開始 >= lng_終了 Then Exit Sub
    
    ' ピボット値（中央の要素を使用）
    var_ピボット = arr_データ((lng_開始 + lng_終了) \ 2, lng_ソート列)
    
    lng_左 = lng_開始
    lng_右 = lng_終了
    
    ' 分割
    Do
        ' 左側からピボットより大きい要素を検索
        Do While IsNumeric(arr_データ(lng_左, lng_ソート列)) And _
                  IsNumeric(var_ピボット) And _
                  CDbl(arr_データ(lng_左, lng_ソート列)) < CDbl(var_ピボット)
            lng_左 = lng_左 + 1
            If lng_左 > lng_終了 Then Exit Do
        Loop
        
        ' 右側からピボットより小さい要素を検索
        Do While IsNumeric(arr_データ(lng_右, lng_ソート列)) And _
                  IsNumeric(var_ピボット) And _
                  CDbl(arr_データ(lng_右, lng_ソート列)) > CDbl(var_ピボット)
            lng_右 = lng_右 - 1
            If lng_右 < lng_開始 Then Exit Do
        Loop
        
        ' 交換
        If lng_左 <= lng_右 Then
            ' 行交換
            For lng_列 = LBound(arr_データ, 2) To UBound(arr_データ, 2)
                var_一時 = arr_データ(lng_左, lng_列)
                arr_データ(lng_左, lng_列) = arr_データ(lng_右, lng_列)
                arr_データ(lng_右, lng_列) = var_一時
            Next lng_列
            
            lng_左 = lng_左 + 1
            lng_右 = lng_右 - 1
        End If
    Loop Until lng_左 > lng_右
    
    ' 再帰的にソート
    If lng_開始 < lng_右 Then クイックソート配列 arr_データ, lng_ソート列, lng_開始, lng_右
    If lng_左 < lng_終了 Then クイックソート配列 arr_データ, lng_ソート列, lng_左, lng_終了
End Sub

'''---------------------------------------------------------
' 5. 重複データの検出と処理
'''---------------------------------------------------------
Sub 重複データ検出処理()
    ' 変数宣言
    Dim ws_データ As Worksheet
    Dim ws_結果 As Worksheet
    Dim rng_データ範囲 As Range
    Dim rng_一意データ As Range
    Dim rng_重複データ As Range
    Dim lng_最終行 As Long
    Dim dict_キー As Object
    Dim var_データ配列 As Variant
    Dim str_キー As String
    Dim lng_行 As Long
    Dim arr_重複インデックス() As Long
    Dim lng_重複数 As Long
    Dim lng_一意数 As Long
    
    ' 画面更新を停止（高速化）
    Application.ScreenUpdating = False
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("データ")
    
    ' 結果シートの作成または取得
    On Error Resume Next
    Set ws_結果 = ThisWorkbook.Worksheets("重複チェック結果")
    If ws_結果 Is Nothing Then
        Set ws_結果 = ThisWorkbook.Worksheets.Add(After:=ws_データ)
        ws_結果.Name = "重複チェック結果"
    Else
        ws_結果.Cells.Clear
    End If
    On Error GoTo 0
    
    ' データ範囲の取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    If lng_最終行 <= 1 Then
        MsgBox "データが見つかりません。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Set rng_データ範囲 = ws_データ.Range("A1:A" & lng_最終行)
    
    ' 1. Excelの組み込み機能を使用した重複データの強調表示
    ' 条件付き書式を適用
    rng_データ範囲.FormatConditions.Delete
    rng_データ範囲.FormatConditions.Add Type:=xlDuplicate
    With rng_データ範囲.FormatConditions(1)
        .Interior.Color = RGB(255, 200, 200)  ' 薄い赤色
        .Font.Color = RGB(156, 0, 6)         ' 濃い赤色
        .Font.Bold = True
    End With
    
    ' 2. VBAを使用した詳細な重複チェック
    ' Dictionaryオブジェクトの作成
    Set dict_キー = CreateObject("Scripting.Dictionary")
    dict_キー.CompareMode = vbTextCompare  ' 大文字小文字を区別しない
    
    ' データを配列に格納
    var_データ配列 = rng_データ範囲.Value
    
    ' 重複インデックス配列の初期化
    ReDim arr_重複インデックス(1 To lng_最終行)
    lng_重複数 = 0
    
    ' 重複チェック
    For lng_行 = 2 To lng_最終行  ' ヘッダー行をスキップ
        str_キー = CStr(var_データ配列(lng_行, 1))
        
        If dict_キー.Exists(str_キー) Then
            ' 重複を見つけた
            lng_重複数 = lng_重複数 + 1
            arr_重複インデックス(lng_重複数) = lng_行
        Else
            ' 新しいキーを追加
            dict_キー.Add str_キー, lng_行
        End If
    Next lng_行
    
    ' 一意のデータ数
    lng_一意数 = dict_キー.Count
    
    ' 3. 重複データの抽出と表示
    ' 結果シートにヘッダーを設定
    ws_結果.Range("A1").Value = "元の行番号"
    ws_結果.Range("B1").Value = "値"
    ws_結果.Range("C1").Value = "ステータス"
    ws_結果.Range("A1:C1").Font.Bold = True
    
    ' 重複データをコピー
    If lng_重複数 > 0 Then
        For lng_行 = 1 To lng_重複数
            ws_結果.Cells(lng_行 + 1, 1).Value = arr_重複インデックス(lng_行)
            ws_結果.Cells(lng_行 + 1, 2).Value = var_データ配列(arr_重複インデックス(lng_行), 1)
            ws_結果.Cells(lng_行 + 1, 3).Value = "重複"
        Next lng_行
    End If
    
    ' 4. 一意データの別シートへのコピー
    If lng_一意数 > 0 Then
        Dim ws_一意 As Worksheet
        
        On Error Resume Next
        Set ws_一意 = ThisWorkbook.Worksheets("一意データ")
        If ws_一意 Is Nothing Then
            Set ws_一意 = ThisWorkbook.Worksheets.Add(After:=ws_結果)
            ws_一意.Name = "一意データ"
        Else
            ws_一意.Cells.Clear
        End If
        On Error GoTo 0
        
        ' ヘッダーのコピー
        ws_一意.Range("A1").Value = var_データ配列(1, 1)
        
        ' 一意データのコピー
        Dim var_キー As Variant
        Dim lng_出力行 As Long
        
        lng_出力行 = 2
        For Each var_キー In dict_キー.Keys
            ws_一意.Cells(lng_出力行, 1).Value = var_キー
            lng_出力行 = lng_出力行 + 1
        Next var_キー
    End If
    
    ' 5. 結果の要約
    Dim ws_サマリ As Worksheet
    
    On Error Resume Next
    Set ws_サマリ = ThisWorkbook.Worksheets("重複チェックサマリ")
    If ws_サマリ Is Nothing Then
        Set ws_サマリ = ThisWorkbook.Worksheets.Add(After:=ws_データ)
        ws_サマリ.Name = "重複チェックサマリ"
    Else
        ws_サマリ.Cells.Clear
    End If
    On Error GoTo 0
    
    ' サマリデータの作成
    ws_サマリ.Range("A1").Value = "重複チェック結果サマリ"
    ws_サマリ.Range("A1").Font.Size = 14
    ws_サマリ.Range("A1").Font.Bold = True
    
    ws_サマリ.Range("A3").Value = "総データ数"
    ws_サマリ.Range("B3").Value = lng_最終行 - 1  ' ヘッダーを除く
    
    ws_サマリ.Range("A4").Value = "一意データ数"
    ws_サマリ.Range("B4").Value = lng_一意数
    
    ws_サマリ.Range("A5").Value = "重複データ数"
    ws_サマリ.Range("B5").Value = lng_重複数
    
    ws_サマリ.Range("A6").Value = "重複率"
    ws_サマリ.Range("B6").Value = Format(lng_重複数 / (lng_最終行 - 1), "0.0%")
    
    ' 書式設定
    ws_サマリ.Columns("A:B").AutoFit
    ws_サマリ.Range("A3:B6").BorderAround ColorIndex:=1
    ws_サマリ.Range("A3:A6").Font.Bold = True
    
    ' クリーンアップ
    Set dict_キー = Nothing
    
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    
    ' サマリシートをアクティブにする
    ws_サマリ.Activate
    
    MsgBox "重複チェックが完了しました。" & vbCrLf & _
           "総データ数: " & lng_最終行 - 1 & vbCrLf & _
           "一意データ数: " & lng_一意数 & vbCrLf & _
           "重複データ数: " & lng_重複数, vbInformation
End Sub


'''---------------------------------------------------------
' 1. Access DBへの接続と基本的なCRUD操作
'''---------------------------------------------------------
Sub AccessDB基本操作()
    ' 変数宣言
    Dim str_DBパス As String
    Dim obj_接続 As Object ' ADODB.Connection
    Dim obj_レコードセット As Object ' ADODB.Recordset
    Dim str_SQL As String
    Dim lng_行 As Long
    Dim ws_データ As Worksheet
    
    ' 参照設定の確認
    ' VBEのツール→参照設定から「Microsoft ActiveX Data Objects x.x Library」を追加している場合は、
    ' 以下のように明示的に型宣言ができます。
    ' Dim obj_接続 As ADODB.Connection
    ' Dim obj_レコードセット As ADODB.Recordset
    ' Set obj_接続 = New ADODB.Connection
    ' Set obj_レコードセット = New ADODB.Recordset
    
    ' 接続先AccessファイルのパスをThisWorkbookと同じフォルダに設定
    str_DBパス = ThisWorkbook.Path & "\データベース.accdb"
    
    ' Accessファイルの存在チェック
    If Dir(str_DBパス) = "" Then
        MsgBox "データベースファイルが見つかりません: " & str_DBパス, vbExclamation
        Exit Sub
    End If
    
    ' ワークシートの設定
    Set ws_データ = ThisWorkbook.Worksheets("DBデータ")
    ws_データ.Cells.Clear
    
    On Error GoTo ErrorHandler
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    Set obj_レコードセット = CreateObject("ADODB.Recordset")
    
    ' 接続文字列の設定（Access 2007以降）
    obj_接続.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & str_DBパス & ";Persist Security Info=False;"
    
    ' データベースに接続
    obj_接続.Open
    
    ' ====== 1. SELECT操作（データ読み取り） ======
    ' SQLクエリの作成
    str_SQL = "SELECT 顧客ID, 顧客名, 住所, 電話番号 FROM 顧客マスタ ORDER BY 顧客ID;"
    
    ' レコードセットを開く
    obj_レコードセット.Open str_SQL, obj_接続, 3, 3 ' CursorType=3:Static, LockType=3:Optimistic
    
    ' データが存在するか確認
    If Not obj_レコードセット.EOF Then
        ' ヘッダー行の設定
        For i = 0 To obj_レコードセット.Fields.Count - 1
            ws_データ.Cells(1, i + 1).Value = obj_レコードセット.Fields(i).Name
        Next i
        
        ' データの設定
        ws_データ.Range("A2").CopyFromRecordset obj_レコードセット
        
        ' 書式設定
        ws_データ.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        ws_データ.Range("A1:D1").Font.Bold = True
        ws_データ.Columns("A:D").AutoFit
    Else
        MsgBox "データが見つかりませんでした。", vbInformation
    End If
    
    ' レコードセットを閉じる
    obj_レコードセット.Close
    
    ' ====== 2. INSERT操作（データ追加） ======
    str_SQL = "INSERT INTO 顧客マスタ (顧客ID, 顧客名, 住所, 電話番号) VALUES (?, ?, ?, ?);"
    
    ' コマンドオブジェクトを使用してパラメータ化クエリを実行
    Dim obj_コマンド As Object ' ADODB.Command
    Set obj_コマンド = CreateObject("ADODB.Command")
    
    With obj_コマンド
        .ActiveConnection = obj_接続
        .CommandText = str_SQL
        .CommandType = 1 ' adCmdText
        
        ' パラメータの追加
        .Parameters.Append .CreateParameter("顧客ID", 200, 1, 10, "C00123") ' adVarChar, adParamInput
        .Parameters.Append .CreateParameter("顧客名", 200, 1, 50, "東京電力 太郎")
        .Parameters.Append .CreateParameter("住所", 200, 1, 100, "東京都千代田区大手町...")
        .Parameters.Append .CreateParameter("電話番号", 200, 1, 20, "03-1234-5678")
        
        ' クエリ実行
        .Execute
    End With
    
    ' ====== 3. UPDATE操作（データ更新） ======
    str_SQL = "UPDATE 顧客マスタ SET 住所 = ?, 電話番号 = ? WHERE 顧客ID = ?;"
    
    ' コマンドオブジェクトを再設定
    Set obj_コマンド = CreateObject("ADODB.Command")
    
    With obj_コマンド
        .ActiveConnection = obj_接続
        .CommandText = str_SQL
        .CommandType = 1 ' adCmdText
        
        ' パラメータの追加
        .Parameters.Append .CreateParameter("住所", 200, 1, 100, "東京都中央区銀座...")
        .Parameters.Append .CreateParameter("電話番号", 200, 1, 20, "03-9876-5432")
        .Parameters.Append .CreateParameter("顧客ID", 200, 1, 10, "C00123")
        
        ' クエリ実行
        .Execute
    End With
    
    ' ====== 4. DELETE操作（データ削除） ======
    str_SQL = "DELETE FROM 顧客マスタ WHERE 顧客ID = ?;"
    
    ' コマンドオブジェクトを再設定
    Set obj_コマンド = CreateObject("ADODB.Command")
    
    With obj_コマンド
        .ActiveConnection = obj_接続
        .CommandText = str_SQL
        .CommandType = 1 ' adCmdText
        
        ' パラメータの追加
        .Parameters.Append .CreateParameter("顧客ID", 200, 1, 10, "C00123")
        
        ' クエリ実行
        .Execute
    End With
    
CleanExit:
    ' 接続を閉じる
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_レコードセット = Nothing
    Set obj_コマンド = Nothing
    Set obj_接続 = Nothing
    
    MsgBox "データベース操作が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'''---------------------------------------------------------
' 2. SQL Server/Oracle等への接続と操作
'''---------------------------------------------------------
Sub SQLServer接続操作()
    ' 変数宣言
    Dim obj_接続 As Object ' ADODB.Connection
    Dim obj_レコードセット As Object ' ADODB.Recordset
    Dim str_接続文字列 As String
    Dim str_SQL As String
    Dim ws_データ As Worksheet
    
    ' ワークシートの設定
    Set ws_データ = ThisWorkbook.Worksheets("SQLデータ")
    ws_データ.Cells.Clear
    
    On Error GoTo ErrorHandler
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    Set obj_レコードセット = CreateObject("ADODB.Recordset")
    
    ' 接続文字列の設定
    ' SQL Server
    str_接続文字列 = "Provider=SQLOLEDB;" & _
                   "Data Source=ServerName;" & _
                   "Initial Catalog=DatabaseName;" & _
                   "Integrated Security=SSPI;"
    
    ' Oracle接続の場合
    'str_接続文字列 = "Provider=OraOLEDB.Oracle;" & _
    '               "Data Source=OracleServiceName;" & _
    '               "User ID=username;" & _
    '               "Password=password;"
    
    ' データベースに接続
    obj_接続.Open str_接続文字列
    
    ' SQLクエリの作成
    str_SQL = "SELECT 顧客ID, 顧客名, 契約種別, 契約電力 " & _
              "FROM 契約マスタ " & _
              "WHERE 契約状態 = '有効' " & _
              "ORDER BY 顧客ID;"
              
    ' レコードセットを開く
    obj_レコードセット.Open str_SQL, obj_接続, 3, 3 ' CursorType=3:Static, LockType=3:Optimistic
    
    ' データが存在するか確認
    If Not obj_レコードセット.EOF Then
        ' ヘッダー行の設定
        For i = 0 To obj_レコードセット.Fields.Count - 1
            ws_データ.Cells(1, i + 1).Value = obj_レコードセット.Fields(i).Name
        Next i
        
        ' データの設定
        ws_データ.Range("A2").CopyFromRecordset obj_レコードセット
        
        ' 書式設定
        ws_データ.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        ws_データ.Range("A1:D1").Font.Bold = True
        ws_データ.Columns("A:D").AutoFit
    Else
        MsgBox "データが見つかりませんでした。", vbInformation
    End If
    
    ' レコードセットを閉じる
    obj_レコードセット.Close
    
    ' ストアドプロシージャの実行
    Dim obj_コマンド As Object ' ADODB.Command
    Set obj_コマンド = CreateObject("ADODB.Command")
    
    With obj_コマンド
        .ActiveConnection = obj_接続
        .CommandText = "SP_更新_契約情報"
        .CommandType = 4 ' adCmdStoredProc
        
        ' パラメータの追加
        .Parameters.Append .CreateParameter("@顧客ID", 200, 1, 10, "C00456") ' adVarChar, adParamInput
        .Parameters.Append .CreateParameter("@契約電力", 3, 1, , 100) ' adInteger, adParamInput
        .Parameters.Append .CreateParameter("@更新者", 200, 1, 50, "システム管理者")
        
        ' 出力パラメータの追加
        Dim prm_結果 As Object ' ADODB.Parameter
        Set prm_結果 = .CreateParameter("@RETURN_VALUE", 3, 2) ' adInteger, adParamOutput
        .Parameters.Append prm_結果
        
        ' ストアドプロシージャ実行
        .Execute
        
        ' 出力パラメータの値を取得
        If prm_結果.Value = 0 Then
            MsgBox "ストアドプロシージャが正常に実行されました。", vbInformation
        Else
            MsgBox "ストアドプロシージャの実行中にエラーが発生しました。エラーコード: " & prm_結果.Value, vbExclamation
        End If
    End With
    
CleanExit:
    ' 接続を閉じる
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_レコードセット = Nothing
    Set obj_コマンド = Nothing
    Set obj_接続 = Nothing
    
    MsgBox "SQL Server操作が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'''---------------------------------------------------------
' 3. パラメータ化クエリによるセキュアな実装
'''---------------------------------------------------------
Sub パラメータ化クエリの実装()
    ' 変数宣言
    Dim obj_接続 As Object ' ADODB.Connection
    Dim obj_コマンド As Object ' ADODB.Command
    Dim obj_レコードセット As Object ' ADODB.Recordset
    Dim str_接続文字列 As String
    Dim str_SQL As String
    Dim ws_検索 As Worksheet
    Dim str_検索条件 As String
    
    ' ワークシートの設定
    Set ws_検索 = ThisWorkbook.Worksheets("検索画面")
    
    ' ユーザー入力の取得（検索条件）
    str_検索条件 = ws_検索.Range("B2").Value
    
    If Trim(str_検索条件) = "" Then
        MsgBox "検索条件を入力してください。", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    Set obj_コマンド = CreateObject("ADODB.Command")
    Set obj_レコードセット = CreateObject("ADODB.Recordset")
    
    ' 接続文字列の設定（例：Access）
    str_接続文字列 = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & ThisWorkbook.Path & "\データベース.accdb;" & _
                   "Persist Security Info=False;"
    
    ' データベースに接続
    obj_接続.Open str_接続文字列
    
    ' ===== 安全でない方法（SQL Injection脆弱性あり）=====
    ' 以下のようなコードは使用しないでください！
    ' str_SQL = "SELECT * FROM 顧客マスタ WHERE 顧客名 LIKE '%" & str_検索条件 & "%'"
    ' obj_レコードセット.Open str_SQL, obj_接続
    
    ' ===== 安全な方法（パラメータ化クエリ）=====
    ' コマンドオブジェクトの設定
    With obj_コマンド
        .ActiveConnection = obj_接続
        .CommandText = "SELECT 顧客ID, 顧客名, 住所, 電話番号 FROM 顧客マスタ WHERE 顧客名 LIKE ?"
        .CommandType = 1 ' adCmdText
        
        ' パラメータの追加
        .Parameters.Append .CreateParameter("顧客名", 200, 1, 50, "%" & str_検索条件 & "%")
        
        ' クエリ実行してレコードセットを取得
        Set obj_レコードセット = .Execute
    End With
    
    ' 結果をシートに表示
    ws_検索.Range("A5:D1000").ClearContents ' 既存の結果をクリア
    
    ' ヘッダー行の設定
    ws_検索.Range("A4").Value = "顧客ID"
    ws_検索.Range("B4").Value = "顧客名"
    ws_検索.Range("C4").Value = "住所"
    ws_検索.Range("D4").Value = "電話番号"
    ws_検索.Range("A4:D4").Font.Bold = True
    
    ' データの設定
    If Not obj_レコードセット.EOF Then
        ws_検索.Range("A5").CopyFromRecordset obj_レコードセット
        ws_検索.Range("A4").CurrentRegion.Borders.LineStyle = xlContinuous
        ws_検索.Columns("A:D").AutoFit
    Else
        ws_検索.Range("A5").Value = "検索条件に一致するデータはありません。"
    End If
    
CleanExit:
    ' 接続を閉じる
    On Error Resume Next
    If Not obj_レコードセット Is Nothing Then
        If obj_レコードセット.State = 1 Then ' adStateOpen
            obj_レコードセット.Close
        End If
    End If
    
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_レコードセット = Nothing
    Set obj_コマンド = Nothing
    Set obj_接続 = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' テーブル作成などDDLクエリのサンプル
Sub テーブル構造操作サンプル()
    ' 変数宣言
    Dim obj_接続 As Object ' ADODB.Connection
    Dim str_接続文字列 As String
    Dim str_SQL As String
    
    On Error GoTo ErrorHandler
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    
    ' 接続文字列の設定（例：Access）
    str_接続文字列 = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & ThisWorkbook.Path & "\データベース.accdb;" & _
                   "Persist Security Info=False;"
    
    ' データベースに接続
    obj_接続.Open str_接続文字列
    
    ' テーブル作成クエリ
    str_SQL = "CREATE TABLE 新規テーブル (" & _
              "ID COUNTER PRIMARY KEY, " & _
              "名称 TEXT(50), " & _
              "数量 INTEGER, " & _
              "金額 CURRENCY, " & _
              "登録日 DATETIME);"
              
    ' クエリを実行
    obj_接続.Execute str_SQL
    
    ' テーブル変更クエリ
    str_SQL = "ALTER TABLE 新規テーブル ADD COLUMN 備考 TEXT(200);"
    obj_接続.Execute str_SQL
    
    ' インデックス作成クエリ
    str_SQL = "CREATE INDEX idx_名称 ON 新規テーブル (名称);"
    obj_接続.Execute str_SQL
    
    MsgBox "テーブル構造の操作が完了しました。", vbInformation
    
CleanExit:
    ' 接続を閉じる
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_接続 = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'''---------------------------------------------------------
' 4. 大量データの一括インポート/エクスポート手法
'''---------------------------------------------------------
Sub 大量データ一括処理()
    ' 変数宣言
    Dim obj_接続 As Object ' ADODB.Connection
    Dim obj_レコードセット As Object ' ADODB.Recordset
    Dim str_接続文字列 As String
    Dim str_SQL As String
    Dim ws_データ As Worksheet
    Dim rng_データ範囲 As Range
    Dim lng_最終行 As Long
    Dim lng_開始行 As Long
    Dim lng_処理件数 As Long
    Dim lng_バッチサイズ As Long
    Dim lng_総件数 As Long
    Dim dbl_開始時間 As Double
    Dim bln_トランザクション中 As Boolean
    
    ' パフォーマンス設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ワークシートの設定
    Set ws_データ = ThisWorkbook.Worksheets("インポートデータ")
    
    ' 最終行の取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    
    If lng_最終行 <= 1 Then
        MsgBox "インポートするデータがありません。", vbExclamation
        GoTo CleanExit
    End If
    
    ' データ範囲の設定（ヘッダー行を除く）
    lng_開始行 = 2
    Set rng_データ範囲 = ws_データ.Range(ws_データ.Cells(lng_開始行, 1), ws_データ.Cells(lng_最終行, 4))
    
    ' 処理件数
    lng_総件数 = lng_最終行 - lng_開始行 + 1
    
    ' バッチサイズの設定（一度に処理するレコード数）
    lng_バッチサイズ = 1000
    
    ' 進捗表示用のステータスバーを初期化
    Application.StatusBar = "データベース接続を準備中..."
    
    On Error GoTo ErrorHandler
    
    ' 開始時間を記録
    dbl_開始時間 = Timer
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    
    ' 接続文字列の設定（例：SQL Server）
    str_接続文字列 = "Provider=SQLOLEDB;" & _
                   "Data Source=ServerName;" & _
                   "Initial Catalog=DatabaseName;" & _
                   "Integrated Security=SSPI;"
    
    ' データベースに接続
    obj_接続.Open str_接続文字列
    
    ' ===== 方法1: レコードセットのAddNewメソッドを使用した挿入 =====
    Set obj_レコードセット = CreateObject("ADODB.Recordset")
    
    ' レコードセットを開く
    str_SQL = "SELECT 顧客ID, 顧客名, 住所, 電話番号 FROM 顧客マスタ"
    obj_レコードセット.Open str_SQL, obj_接続, 1, 3 ' CursorType=1:Keyset, LockType=3:Optimistic
    
    ' バッチ処理を開始
    Dim lng_バッチカウント As Long
    lng_バッチカウント = 0
    lng_処理件数 = 0
    
    ' Excelデータを配列に格納（高速化）
    Dim var_データ配列 As Variant
    var_データ配列 = rng_データ範囲.Value
    
    ' トランザクション開始
    obj_接続.BeginTrans
    bln_トランザクション中 = True
    
    ' データを一括挿入
    Dim i As Long
    For i = 1 To UBound(var_データ配列, 1)
        ' 進捗状況を更新
        If i Mod 100 = 0 Then
            Application.StatusBar = "処理中... " & Format(i / lng_総件数, "0%") & " 完了 (" & i & "/" & lng_総件数 & ")"
            DoEvents
        End If
        
        ' 新しいレコードを追加
        obj_レコードセット.AddNew
        obj_レコードセット.Fields("顧客ID").Value = var_データ配列(i, 1)
        obj_レコードセット.Fields("顧客名").Value = var_データ配列(i, 2)
        obj_レコードセット.Fields("住所").Value = var_データ配列(i, 3)
        obj_レコードセット.Fields("電話番号").Value = var_データ配列(i, 4)
        
        ' バッチカウントを増やす
        lng_バッチカウント = lng_バッチカウント + 1
        
        ' バッチサイズに達したら更新
        If lng_バッチカウント >= lng_バッチサイズ Then
            obj_レコードセット.UpdateBatch
            lng_処理件数 = lng_処理件数 + lng_バッチカウント
            lng_バッチカウント = 0
        End If
    Next i
    
    ' 残りのレコードを更新
    If lng_バッチカウント > 0 Then
        obj_レコードセット.UpdateBatch
        lng_処理件数 = lng_処理件数 + lng_バッチカウント
    End If
    
    ' トランザクションをコミット
    obj_接続.CommitTrans
    bln_トランザクション中 = False
    
    ' レコードセットを閉じる
    obj_レコードセット.Close
    
    ' ===== 方法2: 一括挿入のためのテンポラリテーブルとSQLを使用 =====
    ' テンポラリテーブルを作成（方法1の代わりに使用する場合）
    'str_SQL = "CREATE TABLE #TempCustomers (" & _
    '          "顧客ID VARCHAR(10), " & _
    '          "顧客名 NVARCHAR(50), " & _
    '          "住所 NVARCHAR(100), " & _
    '          "電話番号 VARCHAR(20))"
    'obj_接続.Execute str_SQL
    
    ' BCP（一括コピー）やSQLBulkCopyを使用する場合は、ADO.NETやSQLコマンドラインツールを使用
    
    ' テンポラリテーブルからメインテーブルへのINSERT
    'str_SQL = "INSERT INTO 顧客マスタ (顧客ID, 顧客名, 住所, 電話番号) " & _
    '          "SELECT 顧客ID, 顧客名, 住所, 電話番号 FROM #TempCustomers"
    'obj_接続.Execute str_SQL
    
CleanExit:
    ' トランザクションが開いていれば、ロールバック
    If bln_トランザクション中 Then
        On Error Resume Next
        obj_接続.RollbackTrans
    End If
    
    ' 接続を閉じる
    On Error Resume Next
    If Not obj_レコードセット Is Nothing Then
        If obj_レコードセット.State = 1 Then ' adStateOpen
            obj_レコードセット.Close
        End If
    End If
    
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_レコードセット = Nothing
    Set obj_接続 = Nothing
    
    ' パフォーマンス設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 結果を表示
    If lng_処理件数 > 0 Then
        MsgBox "データの一括処理が完了しました。" & vbCrLf & _
               "処理件数: " & lng_処理件数 & " 件" & vbCrLf & _
               "処理時間: " & Format(Timer - dbl_開始時間, "0.00") & " 秒", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    ' エラー情報を表示
    MsgBox "エラーが発生しました: " & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "説明: " & Err.Description, vbCritical
           
    Resume CleanExit
End Sub

'''---------------------------------------------------------
' 5. 接続エラーのハンドリングと再試行ロジック
'''---------------------------------------------------------
Sub 接続エラーハンドリング()
    ' 変数宣言
    Dim obj_接続 As Object ' ADODB.Connection
    Dim obj_レコードセット As Object ' ADODB.Recordset
    Dim str_接続文字列 As String
    Dim str_SQL As String
    Dim int_再試行回数 As Integer
    Dim int_最大再試行回数 As Integer
    Dim int_再試行間隔_秒 As Integer
    Dim bln_接続成功 As Boolean
    Dim bln_トランザクション中 As Boolean
    
    ' 再試行設定
    int_最大再試行回数 = 3
    int_再試行間隔_秒 = 2
    
    ' ADOオブジェクトの作成
    Set obj_接続 = CreateObject("ADODB.Connection")
    
    ' 接続文字列の設定（例：SQL Server）
    str_接続文字列 = "Provider=SQLOLEDB;" & _
                   "Data Source=ServerName;" & _
                   "Initial Catalog=DatabaseName;" & _
                   "Integrated Security=SSPI;" & _
                   "Connect Timeout=30;"  ' 接続タイムアウトを30秒に設定
    
    ' ===== 接続処理（再試行ロジック） =====
    bln_接続成功 = False
    int_再試行回数 = 0
    
    Do While Not bln_接続成功 And int_再試行回数 <= int_最大再試行回数
        On Error Resume Next
        
        ' 再試行回数をインクリメント
        int_再試行回数 = int_再試行回数 + 1
        
        If int_再試行回数 > 1 Then
            ' 待機メッセージを表示
            Application.StatusBar = "データベース接続の再試行中... (" & int_再試行回数 & "/" & int_最大再試行回数 & ")"
            
            ' 指定秒数だけ待機
            Application.Wait Now + TimeSerial(0, 0, int_再試行間隔_秒)
        End If
        
        ' 接続を試行
        obj_接続.Open str_接続文字列
        
        ' エラーチェック
        If Err.Number = 0 Then
            bln_接続成功 = True
        Else
            ' 接続エラーのログを記録
            Call エラーログ記録(Err.Number, "DB接続エラー: " & Err.Description, "接続エラーハンドリング")
        End If
        
        On Error GoTo 0
    Loop
    
    ' 接続できなかった場合はエラーメッセージを表示して終了
    If Not bln_接続成功 Then
        MsgBox "データベースへの接続に失敗しました。" & vbCrLf & _
               "しばらくしてから再試行してください。", vbCritical
        GoTo CleanExit
    End If
    
    ' ステータスバーをクリア
    Application.StatusBar = False
    
    ' ===== クエリ実行（トランザクション管理） =====
    On Error GoTo ErrorHandler
    
    ' トランザクション開始
    obj_接続.BeginTrans
    bln_トランザクション中 = True
    
    ' クエリ1の実行
    str_SQL = "UPDATE 契約マスタ SET 契約状態 = '解約' WHERE 顧客ID = 'C00123'"
    obj_接続.Execute str_SQL
    
    ' クエリ2の実行
    str_SQL = "INSERT INTO 契約履歴 (顧客ID, 変更日時, 変更内容) VALUES ('C00123', GETDATE(), '契約解約')"
    obj_接続.Execute str_SQL
    
    ' トランザクションのコミット
    obj_接続.CommitTrans
    bln_トランザクション中 = False
    
    MsgBox "データベース操作が正常に完了しました。", vbInformation
    
CleanExit:
    ' 接続を閉じる
    On Error Resume Next
    
    If bln_トランザクション中 Then
        obj_接続.RollbackTrans
    End If
    
    If Not obj_接続 Is Nothing Then
        If obj_接続.State = 1 Then ' adStateOpen
            obj_接続.Close
        End If
    End If
    
    ' オブジェクトの解放
    Set obj_レコードセット = Nothing
    Set obj_接続 = Nothing
    
    ' ステータスバーをクリア
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー情報の表示
    MsgBox "データベース操作中にエラーが発生しました: " & vbCrLf & _
           Err.Description, vbCritical
    
    ' エラーログを記録
    Call エラーログ記録(Err.Number, "DB操作エラー: " & Err.Description, "接続エラーハンドリング")
    
    Resume CleanExit
End Sub

' エラーログを記録する関数
Private Sub エラーログ記録(lng_エラー番号 As Long, str_エラー内容 As String, str_プロシージャ名 As String)
    Dim str_ログパス As String
    Dim int_ファイル番号 As Integer
    
    On Error Resume Next
    
    ' ログファイルパスの設定
    str_ログパス = ThisWorkbook.Path & "\db_error_log.txt"
    
    ' ファイル番号の取得
    int_ファイル番号 = FreeFile
    
    ' ログファイルに追記
    Open str_ログパス For Append As #int_ファイル番号
    Print #int_ファイル番号, Format(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
                           "プロシージャ: " & str_プロシージャ名 & vbTab & _
                           "エラー番号: " & lng_エラー番号 & vbTab & _
                           "内容: " & str_エラー内容
    Close #int_ファイル番号
End Sub
