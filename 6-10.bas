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
    Dim i As Integer ' 未宣言だった変数を追加
    
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
    Dim i As Integer ' 未宣言変数を追加
    
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
    Dim i As Long ' 未宣言変数を追加
    
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
    
    '
