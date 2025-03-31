'''---------------------------------------------------------
' 1. 複数のCSV/テキストファイル一括処理
'''---------------------------------------------------------
Sub 複数CSVファイル一括処理()
    ' 変数宣言
    Dim str_フォルダパス As String
    Dim str_ファイル名 As String
    Dim str_ファイルパス As String
    Dim str_出力パス As String
    Dim ws_集計 As Worksheet
    Dim ws_データ As Worksheet
    Dim lng_最終行 As Long
    Dim lng_出力行 As Long
    Dim int_ファイル番号 As Integer
    Dim str_行データ As String
    Dim arr_列データ() As String
    Dim i As Long
    Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim lng_処理ファイル数 As Long
    
    ' 画面更新を停止して処理を高速化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダパスの設定
    str_フォルダパス = ThisWorkbook.Path & "\データ\"
    
    ' フォルダが存在しない場合は作成
    If Not obj_FSO.FolderExists(str_フォルダパス) Then
        MsgBox "指定されたフォルダが見つかりません: " & str_フォルダパス, vbExclamation
        GoTo CleanExit
    End If
    
    ' 集計シートの設定
    On Error Resume Next
    Set ws_集計 = ThisWorkbook.Worksheets("CSVデータ集計")
    If ws_集計 Is Nothing Then
        Set ws_集計 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_集計.Name = "CSVデータ集計"
    End If
    ws_集計.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' ヘッダー行の設定
    ws_集計.Range("A1").Value = "ファイル名"
    ws_集計.Range("B1").Value = "処理日時"
    ws_集計.Range("C1").Value = "レコード数"
    ws_集計.Range("D1").Value = "合計金額"
    ws_集計.Range("A1:D1").Font.Bold = True
    
    ' フォルダ内のファイルを処理
    Set obj_フォルダ = obj_FSO.GetFolder(str_フォルダパス)
    lng_出力行 = 2
    lng_処理ファイル数 = 0
    
    ' 各CSVファイルを処理
    For Each obj_ファイル In obj_フォルダ.Files
        ' CSVファイルのみを処理
        If LCase(obj_FSO.GetExtensionName(obj_ファイル.Name)) = "csv" Then
            str_ファイル名 = obj_ファイル.Name
            str_ファイルパス = obj_ファイル.Path
            
            ' 処理状況をステータスバーに表示
            Application.StatusBar = "処理中: " & str_ファイル名
            
            ' 一時的なデータシートを準備
            On Error Resume Next
            Set ws_データ = ThisWorkbook.Worksheets("CSVデータ")
            If ws_データ Is Nothing Then
                Set ws_データ = ThisWorkbook.Worksheets.Add(After:=ws_集計)
                ws_データ.Name = "CSVデータ"
            End If
            ws_データ.Cells.Clear
            ws_データ.Visible = xlSheetVeryHidden  ' シートを非表示に
            On Error GoTo ErrorHandler
            
            ' ===== 方法1: FileSystemObjectを使用したCSV読み込み =====
            Dim lng_行数 As Long
            Dim dbl_合計金額 As Double
            
            lng_行数 = 0
            dbl_合計金額 = 0
            
            ' CSVファイルをテキストとして開く
            int_ファイル番号 = FreeFile
            Open str_ファイルパス For Input As #int_ファイル番号
            
            ' 1行目（ヘッダー）を読み込み
            Line Input #int_ファイル番号, str_行データ
            
            ' ヘッダー行をセットアップ
            arr_列データ = Split(str_行データ, ",")
            For i = 0 To UBound(arr_列データ)
                ws_データ.Cells(1, i + 1).Value = arr_列データ(i)
            Next i
            
            ' データ行を読み込み
            lng_行数 = 0
            Do Until EOF(int_ファイル番号)
                Line Input #int_ファイル番号, str_行データ
                lng_行数 = lng_行数 + 1
                
                ' カンマで分割
                arr_列データ = Split(str_行データ, ",")
                
                ' セルに値をセット
                For i = 0 To UBound(arr_列データ)
                    ws_データ.Cells(lng_行数 + 1, i + 1).Value = arr_列データ(i)
                Next i
                
                ' 金額列の値を集計（例: 3列目が金額と仮定）
                If UBound(arr_列データ) >= 2 Then
                    If IsNumeric(arr_列データ(2)) Then
                        dbl_合計金額 = dbl_合計金額 + CDbl(arr_列データ(2))
                    End If
                End If
            Loop
            
            ' ファイルを閉じる
            Close #int_ファイル番号
            
            ' ===== 方法2: QueryTablesを使用したCSV読み込み =====
            ' 以下はQueryTablesを使用した代替手法です（コメントアウト中）
            '
            ' Dim qt As QueryTable
            ' Set qt = ws_データ.QueryTables.Add( _
            '     Connection:="TEXT;" & str_ファイルパス, _
            '     Destination:=ws_データ.Range("A1"))
            '
            ' With qt
            '     .TextFileParseType = xlDelimited
            '     .TextFileCommaDelimiter = True
            '     .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1)  ' 必要に応じて調整
            '     .Refresh BackgroundQuery:=False
            ' End With
            '
            ' ' 金額を集計（例: C列が金額と仮定）
            ' lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
            ' For i = 2 To lng_最終行  ' ヘッダー行をスキップ
            '     If IsNumeric(ws_データ.Cells(i, 3).Value) Then
            '         dbl_合計金額 = dbl_合計金額 + CDbl(ws_データ.Cells(i, 3).Value)
            '     End If
            ' Next i
            
            ' 集計結果をメインシートに記録
            ws_集計.Cells(lng_出力行, 1).Value = str_ファイル名
            ws_集計.Cells(lng_出力行, 2).Value = Now
            ws_集計.Cells(lng_出力行, 3).Value = lng_行数
            ws_集計.Cells(lng_出力行, 4).Value = dbl_合計金額
            
            lng_出力行 = lng_出力行 + 1
            lng_処理ファイル数 = lng_処理ファイル数 + 1
        End If
    Next obj_ファイル
    
    ' 書式設定
    ws_集計.Range("B2:B" & lng_出力行 - 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws_集計.Range("D2:D" & lng_出力行 - 1).NumberFormat = "#,##0.00"
    ws_集計.Columns("A:D").AutoFit
    ws_集計.Activate
    
    ' 処理結果を表示
    MsgBox "CSV一括処理が完了しました。" & vbCrLf & _
           "処理ファイル数: " & lng_処理ファイル数 & " ファイル", vbInformation
    
CleanExit:
    ' 設定を元に戻す
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & _
           Err.Description, vbCritical
    Resume CleanExit
End Sub

'''---------------------------------------------------------
' 2. エクセルファイルの一括変換・統合
'''---------------------------------------------------------
Sub エクセルファイル一括変換統合()
    ' 変数宣言
    Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim str_フォルダパス As String
    Dim str_出力フォルダ As String
    Dim str_出力ブックパス As String
    Dim wb_元 As Workbook
    Dim wb_統合 As Workbook
    Dim ws_元 As Worksheet
    Dim ws_統合 As Worksheet
    Dim ws_目次 As Worksheet
    Dim str_シート名 As String
    Dim lng_処理ファイル数 As Long
    Dim lng_目次行 As Long
    
    ' 画面更新を停止して処理を高速化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダパスの設定
    str_フォルダパス = ThisWorkbook.Path & "\Excelファイル\"
    
    ' 出力フォルダを設定
    str_出力フォルダ = ThisWorkbook.Path & "\出力\"
    
    ' フォルダが存在しない場合は作成
    If Not obj_FSO.FolderExists(str_フォルダパス) Then
        MsgBox "指定されたフォルダが見つかりません: " & str_フォルダパス, vbExclamation
        GoTo CleanExit
    End If
    
    If Not obj_FSO.FolderExists(str_出力フォルダ) Then
        obj_FSO.CreateFolder str_出力フォルダ
    End If
    
    ' 統合ブックを作成
    Set wb_統合 = Workbooks.Add
    
    ' 目次シートを追加
    Set ws_目次 = wb_統合.Sheets(1)
    ws_目次.Name = "目次"
    
    ' 目次シートの設定
    ws_目次.Range("A1").Value = "ファイル一覧"
    ws_目次.Range("A1").Font.Size = 14
    ws_目次.Range("A1").Font.Bold = True
    
    ws_目次.Range("A3").Value = "No."
    ws_目次.Range("B3").Value = "ファイル名"
    ws_目次.Range("C3").Value = "シート名"
    ws_目次.Range("D3").Value = "作成日時"
    ws_目次.Range("A3:D3").Font.Bold = True
    
    lng_目次行 = 4
    lng_処理ファイル数 = 0
    
    ' フォルダ内のファイルを処理
    Set obj_フォルダ = obj_FSO.GetFolder(str_フォルダパス)
    
    For Each obj_ファイル In obj_フォルダ.Files
        ' Excelファイルのみを処理（拡張子のチェック）
        If obj_FSO.GetExtensionName(obj_ファイル.Name) = "xlsx" Or _
           obj_FSO.GetExtensionName(obj_ファイル.Name) = "xls" Then
            
            ' 処理状況をステータスバーに表示
            Application.StatusBar = "処理中: " & obj_ファイル.Name
            
            ' Excelファイルを開く
            On Error Resume Next
            Set wb_元 = Workbooks.Open(obj_ファイル.Path, ReadOnly:=True)
            If Err.Number <> 0 Then
                MsgBox "ファイルを開けませんでした: " & obj_ファイル.Name & vbCrLf & _
                       "エラー: " & Err.Description, vbExclamation
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo ErrorHandler
            
            ' 各シートを処理
            For Each ws_元 In wb_元.Sheets
                ' シート名の取得（特殊文字を除去）
                str_シート名 = クリーンシート名(obj_ファイル.Name & "_" & ws_元.Name)
                
                ' 統合ブックに新しいシートを追加
                Set ws_統合 = wb_統合.Sheets.Add(After:=wb_統合.Sheets(wb_統合.Sheets.Count))
                ws_統合.Name = str_シート名
                
                ' 元シートの全内容をコピー
                ws_元.Cells.Copy ws_統合.Cells
                
                ' ソースファイル情報をシートに追加
                ws_統合.Range("A1").EntireRow.Insert
                ws_統合.Range("A1").Value = "ソース: " & obj_ファイル.Name
                ws_統合.Range("A1").Font.Bold = True
                
                ' 目次を更新
                ws_目次.Range("A" & lng_目次行).Value = lng_処理ファイル数 + 1
                ws_目次.Range("B" & lng_目次行).Value = obj_ファイル.Name
                ws_目次.Range("C" & lng_目次行).Value = ws_元.Name
                ws_目次.Range("D" & lng_目次行).Value = obj_ファイル.DateLastModified
                
                ' ハイパーリンクの追加
                ws_目次.Hyperlinks.Add _
                    Anchor:=ws_目次.Range("C" & lng_目次行), _
                    Address:="", _
                    SubAddress:="'" & str_シート名 & "'!A1", _
                    TextToDisplay:=ws_元.Name
                
                lng_目次行 = lng_目次行 + 1
            Next ws_元
            
            ' 元ブックを閉じる
            wb_元.Close SaveChanges:=False
            
            lng_処理ファイル数 = lng_処理ファイル数 + 1
        End If
NextFile:
    Next obj_ファイル
    
    ' 目次の書式設定
    ws_目次.Columns("A:D").AutoFit
    ws_目次.Range("A3:D" & lng_目次行 - 1).Borders.LineStyle = xlContinuous
    ws_目次.Range("D4:D" & lng_目次行 - 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    
    ' 目次シートをアクティブにする
    ws_目次.Activate
    ws_目次.Range("A1").Select
    
    ' 統合ブックを保存
    str_出力ブックパス = str_出力フォルダ & "統合ブック_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    wb_統合.SaveAs str_出力ブックパス
    
    ' 処理結果を表示
    MsgBox "Excel一括変換・統合が完了しました。" & vbCrLf & _
           "処理ファイル数: " & lng_処理ファイル数 & " ファイル" & vbCrLf & _
           "出力先: " & str_出力ブックパス, vbInformation
    
CleanExit:
    ' 設定を元に戻す
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & _
           Err.Description, vbCritical
    Resume CleanExit
End Sub

' シート名を有効な形式にクリーンアップする関数
Private Function クリーンシート名(str_元の名前 As String) As String
    Dim str_結果 As String
    Dim lng_最大長 As Long
    Dim i As Integer
    Dim str_禁止文字 As String
    
    ' シート名の最大長（31文字）
    lng_最大長 = 31
    
    ' 禁止文字の設定
    str_禁止文字 = "\/[]?*:"
    
    ' まずは長さを制限
    If Len(str_元の名前) > lng_最大長 Then
        str_結果 = Left(str_元の名前, lng_最大長)
    Else
        str_結果 = str_元の名前
    End If
    
    ' 禁止文字を置換
    For i = 1 To Len(str_禁止文字)
        str_結果 = Replace(str_結果, Mid(str_禁止文字, i, 1), "_")
    Next i
    
    クリーンシート名 = str_結果
End Function

'''---------------------------------------------------------
' 3. ネットワークパス・共有フォルダ操作の安定化
'''---------------------------------------------------------
Sub ネットワークフォルダ操作安定化()
    ' 変数宣言
    Dim str_ネットワークパス As String
    Dim str_ローカルパス As String
    Dim str_ドライブ文字 As String
    Dim obj_Shell As Object
    Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim str_ファイル名 As String
    Dim bln_接続成功 As Boolean
    Dim int_再試行回数 As Integer
    Dim int_最大再試行回数 As Integer
    
    ' 再試行設定
    int_最大再試行回数 = 3
    
    ' シェルオブジェクトとFSOの作成
    Set obj_Shell = CreateObject("WScript.Shell")
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' ネットワークパスとマップするドライブの設定
    str_ネットワークパス = "\\サーバー名\共有フォルダ\データ"
    str_ドライブ文字 = "Z:"
    
    ' ===== 1. ネットワークドライブのマッピング =====
    bln_接続成功 = False
    int_再試行回数 = 0
    
    Do While Not bln_接続成功 And int_再試行回数 < int_最大再試行回数
        int_再試行回数 = int_再試行回数 + 1
        
        On Error Resume Next
        
        ' 既存のマッピングを削除（念のため）
        obj_Shell.Run "net use " & str_ドライブ文字 & " /delete /y", 0, True
        
        ' ネットワークドライブをマップ
        ' ※資格情報が必要な場合は /user:ドメイン\ユーザー名 パスワード を追加
        obj_Shell.Run "net use " & str_ドライブ文字 & " """ & str_ネットワークパス & """", 0, True
        
        ' 接続結果の確認
        If Err.Number = 0 And obj_FSO.DriveExists(Left(str_ドライブ文字, 1)) Then
            bln_接続成功 = True
        Else
            ' 待機してから再試行
            Application.Wait Now + TimeSerial(0, 0, 3)
        End If
        
        On Error GoTo ErrorHandler
    Loop
    
    ' 接続に失敗した場合は終了
    If Not bln_接続成功 Then
        MsgBox "ネットワークドライブへの接続に失敗しました。" & vbCrLf & _
               "ネットワーク接続を確認してください。", vbCritical
        GoTo CleanupDrive
    End If
    
    ' マップしたドライブのパスを設定
    str_ローカルパス = str_ドライブ文字 & "\"
    
    ' ===== 2. ファイル操作の安定化 =====
    On Error Resume Next
    
    ' ファイルが存在するか確認してからコピー
    If obj_FSO.FileExists(str_ローカルパス & "ソースファイル.xlsx") Then
        obj_FSO.CopyFile str_ローカルパス & "ソースファイル.xlsx", ThisWorkbook.Path & "\", True
        
        If Err.Number <> 0 Then
            MsgBox "ファイルのコピーに失敗しました: " & Err.Description, vbExclamation
            Err.Clear
        End If
    Else
        MsgBox "ソースファイルが見つかりません: " & str_ローカルパス & "ソースファイル.xlsx", vbExclamation
    End If
    
    ' ===== 3. ファイル一覧取得と処理 =====
    On Error Resume Next
    
    If obj_FSO.FolderExists(str_ローカルパス) Then
        Set obj_フォルダ = obj_FSO.GetFolder(str_ローカルパス)
        
        ' 結果表示用の文字列
        Dim str_ファイル一覧 As String
        str_ファイル一覧 = "ネットワークフォルダ内のファイル一覧:" & vbCrLf & vbCrLf
        
        ' フォルダ内のファイルを列挙
        For Each obj_ファイル In obj_フォルダ.Files
            str_ファイル一覧 = str_ファイル一覧 & obj_ファイル.Name & vbCrLf
        Next obj_ファイル
        
        ' 結果を表示
        MsgBox str_ファイル一覧, vbInformation, "ネットワークフォルダ内ファイル一覧"
    Else
        MsgBox "指定されたネットワークフォルダが見つかりません: " & str_ローカルパス, vbExclamation
    End If
    
    On Error GoTo ErrorHandler
    
    ' ===== 4. ファイルコピーのバッチ処理 =====
    ' （実際のアプリケーションに応じて実装）
    
CleanupDrive:
    ' ネットワークドライブの切断
    On Error Resume Next
    obj_Shell.Run "net use " & str_ドライブ文字 & " /delete /y", 0, True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & vbCrLf & _
           Err.Description, vbCritical
    
    ' ネットワークドライブを切断
    On Error Resume Next
    obj_Shell.Run "net use " & str_ドライブ文字 & " /delete /y", 0, True
End Sub

'''---------------------------------------------------------
' 4. ログファイル生成と管理の自動化
'''---------------------------------------------------------
Sub ログファイル管理自動化()
    ' 変数宣言
    Dim str_ログフォルダ As String
    Dim str_ログファイル As String
    Dim obj_FSO As Object
    Dim obj_ファイル As Object
    Dim obj_フォルダ As Object
    Dim obj_テキストストリーム As Object
    Dim obj_ログファイル As Object
    Dim str_アプリ名 As String
    Dim str_ユーザー名 As String
    Dim dt_現在 As Date
    Dim dt_保存期間制限 As Date
    Dim lng_削除ファイル数 As Long
    
    ' 定数宣言
    Const INT_ログ保存日数 As Integer = 30  ' ログを保持する日数
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' ログフォルダの設定
    str_ログフォルダ = ThisWorkbook.Path & "\logs\"
    
    ' アプリケーション情報の設定
    str_アプリ名 = "電力データ処理システム"
    str_ユーザー名 = Application.UserName
    
    ' 日付の設定
    dt_現在 = Now
    
    ' ===== 1. ログフォルダの確認と作成 =====
    If Not obj_FSO.FolderExists(str_ログフォルダ) Then
        obj_FSO.CreateFolder str_ログフォルダ
    End If
    
    ' ログファイルのオープン/作成
    On Error Resume Next
    If obj_FSO.FileExists(str_ログファイル) Then
        ' 既存のファイルに追記
        Set obj_テキストストリーム = obj_FSO.OpenTextFile(str_ログファイル, 8) ' 8 = 追記モード
    Else
        ' 新規ファイルを作成
        Set obj_テキストストリーム = obj_FSO.CreateTextFile(str_ログファイル, True)
        
        ' ヘッダー情報を記録
        obj_テキストストリーム.WriteLine "===== ログファイル開始: " & Format(Now, "yyyy/mm/dd") & " ====="
        obj_テキストストリーム.WriteLine "アプリケーション: " & ThisWorkbook.Name
        obj_テキストストリーム.WriteLine "ユーザー: " & Application.UserName
        obj_テキストストリーム.WriteLine "========================================"
        obj_テキストストリーム.WriteLine ""
    End If
    On Error GoTo 0
    
    ' タイムスタンプ付きでメッセージを記録
    obj_テキストストリーム.WriteLine Format(Now, "yyyy/mm/dd hh:nn:ss") & " [" & str_レベル & "] " & str_メッセージ
    
    ' ファイルを閉じる
    obj_テキストストリーム.Close
End Sub

'''---------------------------------------------------------
' 5. ファイル名・パスの動的構築と検証
'''---------------------------------------------------------
Sub ファイルパス動的構築検証()
    ' 変数宣言
    Dim str_基本パス As String
    Dim str_環境名 As String
    Dim str_年月 As String
    Dim str_ファイル名 As String
    Dim str_完全パス As String
    Dim arr_環境リスト() As String
    Dim str_環境 As String
    Dim obj_FSO As Object
    Dim i As Long
    Dim str_検証結果 As String
    
    ' 配列の初期化
    arr_環境リスト = Split("開発環境,テスト環境,本番環境", ",")
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' 基本パスの設定
    str_基本パス = ThisWorkbook.Path
    
    ' 年月を取得（YYYYMM形式）
    str_年月 = Format(Date, "yyyymm")
    
    ' 検証結果の初期化
    str_検証結果 = "【パス検証結果】" & vbCrLf & vbCrLf
    
    ' ===== 1. 環境ごとのパス構築と検証 =====
    For i = 0 To UBound(arr_環境リスト)
        str_環境名 = arr_環境リスト(i)
        
        ' ファイル名の構築（例：'契約データ_開発環境_202401.xlsx'）
        str_ファイル名 = "契約データ_" & str_環境名 & "_" & str_年月 & ".xlsx"
        
        ' フォルダパスの構築（環境別フォルダ）
        Dim str_フォルダパス As String
        str_フォルダパス = str_基本パス & "\" & str_環境名 & "\" & str_年月 & "\"
        
        ' 完全パスの構築
        str_完全パス = str_フォルダパス & str_ファイル名
        
        ' パスの検証
        str_検証結果 = str_検証結果 & "環境: " & str_環境名 & vbCrLf
        str_検証結果 = str_検証結果 & "フォルダパス: " & str_フォルダパス & vbCrLf
        
        ' フォルダの存在チェック
        If obj_FSO.FolderExists(str_フォルダパス) Then
            str_検証結果 = str_検証結果 & "フォルダ状態: 存在します" & vbCrLf
        Else
            str_検証結果 = str_検証結果 & "フォルダ状態: 存在しません（作成が必要）" & vbCrLf
            
            ' フォルダの自動作成（オプション）
            On Error Resume Next
            Call フォルダ階層作成(str_フォルダパス)
            If Err.Number = 0 Then
                str_検証結果 = str_検証結果 & "  → フォルダを作成しました" & vbCrLf
            Else
                str_検証結果 = str_検証結果 & "  → フォルダ作成エラー: " & Err.Description & vbCrLf
            End If
            On Error GoTo 0
        End If
        
        ' ファイルの存在チェック
        str_検証結果 = str_検証結果 & "ファイルパス: " & str_完全パス & vbCrLf
        
        If obj_FSO.FileExists(str_完全パス) Then
            str_検証結果 = str_検証結果 & "ファイル状態: 存在します" & vbCrLf
            
            ' ファイル情報の取得
            Dim obj_ファイル As Object
            Set obj_ファイル = obj_FSO.GetFile(str_完全パス)
            
            str_検証結果 = str_検証結果 & "  - サイズ: " & Format(obj_ファイル.Size / 1024, "#,##0.00") & " KB" & vbCrLf
            str_検証結果 = str_検証結果 & "  - 更新日時: " & obj_ファイル.DateLastModified & vbCrLf
        Else
            str_検証結果 = str_検証結果 & "ファイル状態: 存在しません" & vbCrLf
        End If
        
        str_検証結果 = str_検証結果 & vbCrLf
    Next i
    
    ' ===== 2. ファイル名の有効性検証 =====
    str_検証結果 = str_検証結果 & "【ファイル名の有効性検証】" & vbCrLf & vbCrLf
    
    ' 無効な文字を含むファイル名のサンプル
    Dim str_無効なファイル名 As String
    str_無効なファイル名 = "契約データ_<2024/01>.xlsx"
    
    str_検証結果 = str_検証結果 & "検証対象: " & str_無効なファイル名 & vbCrLf
    
    ' ファイル名の修正
    Dim str_有効なファイル名 As String
    str_有効なファイル名 = 有効なファイル名に変換(str_無効なファイル名)
    
    str_検証結果 = str_検証結果 & "修正後: " & str_有効なファイル名 & vbCrLf & vbCrLf
    
    ' ===== 3. パスの長さの検証 =====
    str_検証結果 = str_検証結果 & "【パスの長さの検証】" & vbCrLf & vbCrLf
    
    ' 非常に長いパスの例
    Dim str_長いパス As String
    str_長いパス = str_基本パス & "\非常に長いフォルダ名のサンプル\さらに長いサブフォルダ名\" & _
                   "もっと長いサブフォルダ名\とても長いファイル名の例でWindowsの最大パス長を超えそうな場合.xlsx"
    
    str_検証結果 = str_検証結果 & "長いパス: " & str_長いパス & vbCrLf
    str_検証結果 = str_検証結果 & "パス長: " & Len(str_長いパス) & " 文字" & vbCrLf
    
    ' Windowsの最大パス長（通常は260文字）との比較
    If Len(str_長いパス) > 260 Then
        str_検証結果 = str_検証結果 & "警告: パスがWindowsの最大長（260文字）を超えています。" & vbCrLf
        str_検証結果 = str_検証結果 & "推奨: より短いパスを使用するか、長いパス対応のUNC形式を検討してください。" & vbCrLf
    Else
        str_検証結果 = str_検証結果 & "問題ありません: パス長はWindowsの制限内です。" & vbCrLf
    End If
    
    ' 結果を表示
    MsgBox str_検証結果, vbInformation, "ファイルパス検証結果"
End Sub

' 階層的にフォルダを作成する関数
Private Sub フォルダ階層作成(str_パス As String)
    Dim obj_FSO As Object
    Dim arr_フォルダ As Variant
    Dim str_現在パス As String
    Dim i As Long
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' パスからドライブ部分を取得
    Dim str_ドライブ As String
    If InStr(2, str_パス, ":") = 2 Then
        str_ドライブ = Left(str_パス, 2)
        str_現在パス = str_ドライブ
    Else
        str_現在パス = ""
    End If
    
    ' パスを区切り文字で分割
    arr_フォルダ = Split(Replace(str_パス, str_ドライブ, ""), "\")
    
    ' 各階層のフォルダを順に作成
    For i = 0 To UBound(arr_フォルダ)
        If arr_フォルダ(i) <> "" Then
            str_現在パス = str_現在パス & "\" & arr_フォルダ(i)
            
            If Not obj_FSO.FolderExists(str_現在パス) Then
                obj_FSO.CreateFolder str_現在パス
            End If
        End If
    Next i
End Sub

' ファイル名に含まれる無効な文字を置換する関数
Private Function 有効なファイル名に変換(str_元ファイル名 As String) As String
    Dim str_結果 As String
    Dim str_無効文字 As String
    Dim i As Long
    
    ' ファイル名に使用できない文字のリスト
    str_無効文字 = "\/:*?""<>|"
    
    str_結果 = str_元ファイル名
    
    ' 無効な文字をアンダースコアに置換
    For i = 1 To Len(str_無効文字)
        str_結果 = Replace(str_結果, Mid(str_無効文字, i, 1), "_")
    Next i
    
    有効なファイル名に変換 = str_結果
End Function
    
    ' ===== 2. 新しいログファイルの作成 =====
    ' ログファイル名を日付と時刻で生成
    str_ログファイル = str_ログフォルダ & "Log_" & _
                        Format(dt_現在, "yyyymmdd_hhnnss") & ".txt"
    
    ' ログファイルの作成
    Set obj_テキストストリーム = obj_FSO.CreateTextFile(str_ログファイル, True)
    
    ' ヘッダー情報の書き込み
    With obj_テキストストリーム
        .WriteLine "===== " & str_アプリ名 & " ログファイル ====="
        .WriteLine "日時: " & Format(dt_現在, "yyyy/mm/dd hh:nn:ss")
        .WriteLine "ユーザー: " & str_ユーザー名
        .WriteLine "Excelバージョン: " & Application.Version
        .WriteLine "======================================"
        .WriteLine ""
        
        ' 処理内容のログを記録（サンプル）
        .WriteLine Format(Now, "hh:nn:ss") & " - 処理開始"
        .WriteLine Format(Now, "hh:nn:ss") & " - データファイルの読み込み"
        .WriteLine Format(Now, "hh:nn:ss") & " - レコード数: 1,234件"
        .WriteLine Format(Now, "hh:nn:ss") & " - 処理完了"
        
        ' ログファイルを閉じる
        .Close
    End With
    
    ' ===== 3. 古いログファイルの管理 =====
    ' 保存期間を過ぎたファイルを削除
    dt_保存期間制限 = DateAdd("d", -INT_ログ保存日数, dt_現在)
    lng_削除ファイル数 = 0
    
    Set obj_フォルダ = obj_FSO.GetFolder(str_ログフォルダ)
    
    For Each obj_ファイル In obj_フォルダ.Files
        ' ログファイルのみを対象
        If LCase(obj_FSO.GetExtensionName(obj_ファイル.Name)) = "txt" And _
           Left(obj_ファイル.Name, 4) = "Log_" Then
            
            ' ファイルの作成日が保存期間を超えているか確認
            If obj_ファイル.DateCreated < dt_保存期間制限 Then
                On Error Resume Next
                obj_FSO.DeleteFile obj_ファイル.Path, True
                
                If Err.Number = 0 Then
                    lng_削除ファイル数 = lng_削除ファイル数 + 1
                End If
                On Error GoTo 0
            End If
        End If
    Next obj_ファイル
    
    ' ===== 4. ログサマリーの作成 =====
    ' 最新のログファイルをオープン
    Set obj_テキストストリーム = obj_FSO.OpenTextFile(str_ログファイル, 8) ' 8 = 追記モード
    
    ' ログ管理情報を追記
    With obj_テキストストリーム
        .WriteLine ""
        .WriteLine "===== ログ管理情報 ====="
        .WriteLine "ログフォルダ: " & str_ログフォルダ
        .WriteLine "ログ保存期間: " & INT_ログ保存日数 & " 日"
        .WriteLine "古いログファイル削除数: " & lng_削除ファイル数 & " ファイル"
        .WriteLine "===== ログ終了 ====="
        
        ' ログファイルを閉じる
        .Close
    End With
    
    ' メッセージを表示
    MsgBox "ログファイルが作成されました: " & vbCrLf & _
           str_ログファイル & vbCrLf & vbCrLf & _
           "保存期間（" & INT_ログ保存日数 & "日）を過ぎたログファイル " & _
           lng_削除ファイル数 & " 件を削除しました。", vbInformation
End Sub

' ログファイルに記録する補助関数
Public Sub ログ記録(str_メッセージ As String, Optional str_レベル As String = "INFO")
    ' 変数宣言
    Dim str_ログフォルダ As String
    Dim str_ログファイル As String
    Dim obj_FSO As Object
    Dim obj_テキストストリーム As Object
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' ログフォルダの設定
    str_ログフォルダ = ThisWorkbook.Path & "\logs\"
    
    ' 日付に基づくログファイル名（1日1ファイル）
    str_ログファイル = str_ログフォルダ & "Log_" & Format(Date, "yyyymmdd") & ".txt"
    
    ' ログフォルダの確認と作成
    If Not obj_FSO.FolderExists(str_ログフォルダ) Then
        ' ログフォルダが存在しない場合は作成する
        obj_FSO.CreateFolder str_ログフォルダ
    End If
    
    ' ログファイルを開く（存在しない場合は作成される）
    ' 8 = ForAppending（追記モード）
    Set obj_テキストストリーム = obj_FSO.OpenTextFile(str_ログファイル, 8, True)
    
    ' タイムスタンプとレベルを含めてメッセージを書き込む
    obj_テキストストリーム.WriteLine Format(Now, "yyyy/mm/dd hh:nn:ss") & " [" & str_レベル & "] " & str_メッセージ
    
    ' ファイルを閉じる
    obj_テキストストリーム.Close
    
    ' オブジェクトの解放
    Set obj_テキストストリーム = Nothing
    Set obj_FSO = Nothing
End Sub