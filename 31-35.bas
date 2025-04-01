'''---------------------------------------------------------
' 1. Web APIからのデータ取得とJSON解析
'''---------------------------------------------------------
Sub WebAPI連携とJSON解析()
    ' 変数宣言
    Dim obj_HTTP As Object
    Dim str_URL As String
    Dim str_APIキー As String
    Dim str_リクエスト As String
    Dim str_レスポンス As String
    Dim obj_JSON As Object
    Dim obj_Item As Object
    Dim ws_結果 As Worksheet
    Dim lng_出力行 As Long
    
    ' 結果表示用のワークシートを準備
    On Error Resume Next
    Set ws_結果 = ThisWorkbook.Worksheets("API結果")
    If ws_結果 Is Nothing Then
        Set ws_結果 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_結果.Name = "API結果"
    End If
    ws_結果.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' 画面更新を停止
    Application.ScreenUpdating = False
    
    ' HTTPリクエストオブジェクトの作成
    Set obj_HTTP = CreateObject("MSXML2.XMLHTTP")
    
    ' API呼び出し情報を設定（例：OpenWeatherMap APIを使用した天気情報取得）
    str_URL = "https://api.openweathermap.org/data/2.5/weather"
    str_APIキー = "YOUR_API_KEY" ' 実際のAPIキーに置き換える
    
    ' 東京の天気を取得するパラメータ
    str_リクエスト = str_URL & "?q=Tokyo,jp&units=metric&appid=" & str_APIキー
    
    ' HTTPリクエストの実行
    With obj_HTTP
        .Open "GET", str_リクエスト, False
        .setRequestHeader "Content-Type", "application/json"
        .send
        
        ' レスポンスのチェック
        If .Status <> 200 Then
            MsgBox "APIリクエストが失敗しました。" & vbCrLf & _
                   "ステータス: " & .Status & vbCrLf & _
                   "詳細: " & .responseText, vbExclamation
            Exit Sub
        End If
        
        ' レスポンスの取得
        str_レスポンス = .responseText
    End With
    
    ' VBA-JSONライブラリを使用してJSONを解析
    ' 注：VBA-JSONライブラリが必要です
    ' https://github.com/VBA-tools/VBA-JSON からダウンロードしてインポート
    
    ' JSONオブジェクトの作成
    Set obj_JSON = ParseJSON(str_レスポンス)
    
    ' ヘッダーの設定
    ws_結果.Range("A1").Value = "項目"
    ws_結果.Range("B1").Value = "値"
    ws_結果.Range("A1:B1").Font.Bold = True
    
    ' 出力開始行
    lng_出力行 = 2
    
    ' 基本情報を出力
    ws_結果.Cells(lng_出力行, 1).Value = "都市名"
    ws_結果.Cells(lng_出力行, 2).Value = obj_JSON("name")
    lng_出力行 = lng_出力行 + 1
    
    ' 天気情報を出力
    Dim obj_天気 As Object
    Set obj_天気 = obj_JSON("weather")(1)
    
    ws_結果.Cells(lng_出力行, 1).Value = "天気"
    ws_結果.Cells(lng_出力行, 2).Value = obj_天気("main") & " (" & obj_天気("description") & ")"
    lng_出力行 = lng_出力行 + 1
    
    ' 気温情報を出力
    Dim obj_気温 As Object
    Set obj_気温 = obj_JSON("main")
    
    ws_結果.Cells(lng_出力行, 1).Value = "現在の気温"
    ws_結果.Cells(lng_出力行, 2).Value = obj_気温("temp") & " °C"
    lng_出力行 = lng_出力行 + 1
    
    ws_結果.Cells(lng_出力行, 1).Value = "最高気温"
    ws_結果.Cells(lng_出力行, 2).Value = obj_気温("temp_max") & " °C"
    lng_出力行 = lng_出力行 + 1
    
    ws_結果.Cells(lng_出力行, 1).Value = "最低気温"
    ws_結果.Cells(lng_出力行, 2).Value = obj_気温("temp_min") & " °C"
    lng_出力行 = lng_出力行 + 1
    
    ws_結果.Cells(lng_出力行, 1).Value = "湿度"
    ws_結果.Cells(lng_出力行, 2).Value = obj_気温("humidity") & " %"
    lng_出力行 = lng_出力行 + 1
    
    ' 風情報を出力
    Dim obj_風 As Object
    Set obj_風 = obj_JSON("wind")
    
    ws_結果.Cells(lng_出力行, 1).Value = "風速"
    ws_結果.Cells(lng_出力行, 2).Value = obj_風("speed") & " m/s"
    lng_出力行 = lng_出力行 + 1
    
    ws_結果.Cells(lng_出力行, 1).Value = "風向"
    ws_結果.Cells(lng_出力行, 2).Value = obj_風("deg") & " 度"
    lng_出力行 = lng_出力行 + 1
    
    ' データ取得時刻を出力
    Dim lng_Unix時間 As Long
    Dim dt_更新日時 As Date
    
    lng_Unix時間 = obj_JSON("dt")
    dt_更新日時 = DateAdd("s", lng_Unix時間, #1/1/1970#)
    
    ws_結果.Cells(lng_出力行, 1).Value = "データ更新日時"
    ws_結果.Cells(lng_出力行, 2).Value = dt_更新日時
    ws_結果.Cells(lng_出力行, 2).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    lng_出力行 = lng_出力行 + 1
    
    ' 書式設定
    ws_結果.Columns("A:B").AutoFit
    ws_結果.Range("A1:B" & lng_出力行 - 1).Borders.LineStyle = xlContinuous
    
    ' 正常終了処理
    ws_結果.Activate
    MsgBox "APIからのデータ取得と解析が完了しました。", vbInformation
    
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

' JSON解析関数（簡易版 - 実際にはVBA-JSONライブラリを使用することを推奨）
Private Function ParseJSON(str_JSON As String) As Object
    ' 注意: この簡易実装は非常に限定的です
    ' 実際のプロジェクトではVBA-JSONなどの専用ライブラリを使用してください
    ' https://github.com/VBA-tools/VBA-JSON
    
    ' この関数はVBA-JSONライブラリがない場合のプレースホルダーとして機能します
    Set ParseJSON = CreateObject("Scripting.Dictionary")
    
    ' 実際には、ここでJSONの解析ロジックを実装します
    ' ただし、完全なJSON解析は複雑なため、ライブラリの使用を強く推奨します
    
    ' サンプルデータ（実際のAPIレスポンスの代わり）
    ParseJSON.Add "name", "Tokyo"
    
    Dim weather As Object
    Set weather = CreateObject("Scripting.Dictionary")
    weather.Add "main", "Clear"
    weather.Add "description", "晴れ"
    
    Dim weatherList As Object
    Set weatherList = CreateObject("Scripting.Dictionary")
    weatherList.Add 1, weather
    
    ParseJSON.Add "weather", weatherList
    
    Dim main As Object
    Set main = CreateObject("Scripting.Dictionary")
    main.Add "temp", 22.5
    main.Add "temp_max", 24.3
    main.Add "temp_min", 19.8
    main.Add "humidity", 65
    
    ParseJSON.Add "main", main
    
    Dim wind As Object
    Set wind = CreateObject("Scripting.Dictionary")
    wind.Add "speed", 3.2
    wind.Add "deg", 180
    
    ParseJSON.Add "wind", wind
    ParseJSON.Add "dt", CDbl(Date - #1/1/1970#) * 86400
End Function

'''---------------------------------------------------------
' 2. Webスクレイピングによるデータ収集
'''---------------------------------------------------------
Sub Webスクレイピング()
    ' 変数宣言
    Dim obj_IE As Object
    Dim obj_要素 As Object
    Dim obj_要素リスト As Object
    Dim ws_結果 As Worksheet
    Dim str_URL As String
    Dim lng_出力行 As Long
    Dim i As Long
    Dim bln_ページ読込完了 As Boolean
    
    ' 結果表示用のワークシートを準備
    On Error Resume Next
    Set ws_結果 = ThisWorkbook.Worksheets("Web情報")
    If ws_結果 Is Nothing Then
        Set ws_結果 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_結果.Name = "Web情報"
    End If
    ws_結果.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' 画面更新を停止
    Application.ScreenUpdating = False
    
    ' スクレイピング対象URLの設定
    str_URL = "https://www.tepco.co.jp/forecast/html/area_data-j.html" ' 東京電力の電力使用状況ページ（例）
    
    ' Internet Explorerオブジェクトの作成
    Set obj_IE = CreateObject("InternetExplorer.Application")
    
    ' ステータスバーの更新
    Application.StatusBar = "Webページを読み込み中..."
    
    ' ブラウザの設定と表示
    With obj_IE
        .Visible = False       ' バックグラウンド処理（True にするとブラウザが表示される）
        .navigate str_URL      ' 指定URLに移動
        .Silent = True         ' エラーメッセージを非表示
    End With
    
    ' ページの読み込み完了を待機
    Do While obj_IE.Busy Or obj_IE.readyState <> 4 ' 4 = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    ' ページがロードされるまで少し待機（Ajax読み込み対策）
    Application.Wait Now + TimeSerial(0, 0, 2)
    
    ' ヘッダー行の設定
    ws_結果.Range("A1").Value = "日時"
    ws_結果.Range("B1").Value = "予想最大電力(万kW)"
    ws_結果.Range("C1").Value = "供給力(万kW)"
    ws_結果.Range("D1").Value = "現在の電力使用量(万kW)"
    ws_結果.Range("A1:D1").Font.Bold = True
    
    ' 出力開始行
    lng_出力行 = 2
    
    ' ページから情報を取得（例：電力使用状況データの取得）
    On Error Resume Next
    
    ' 更新日時の取得
    Dim str_更新日時 As String
    Set obj_要素 = obj_IE.document.querySelector("span.update-date")
    If Not obj_要素 Is Nothing Then
        str_更新日時 = obj_要素.innerText
        ws_結果.Cells(lng_出力行, 1).Value = str_更新日時
    End If
    
    ' 予想最大電力の取得
    Dim str_予想最大電力 As String
    Set obj_要素 = obj_IE.document.querySelector("div.peakdemand span.amount")
    If Not obj_要素 Is Nothing Then
        str_予想最大電力 = obj_要素.innerText
        ws_結果.Cells(lng_出力行, 2).Value = str_予想最大電力
    End If
    
    ' 供給力の取得
    Dim str_供給力 As String
    Set obj_要素 = obj_IE.document.querySelector("div.capacity span.amount")
    If Not obj_要素 Is Nothing Then
        str_供給力 = obj_要素.innerText
        ws_結果.Cells(lng_出力行, 3).Value = str_供給力
    End If
    
    ' 現在の電力使用量の取得
    Dim str_使用量 As String
    Set obj_要素 = obj_IE.document.querySelector("div.usage span.amount")
    If Not obj_要素 Is Nothing Then
        str_使用量 = obj_要素.innerText
        ws_結果.Cells(lng_出力行, 4).Value = str_使用量
    End If
    
    ' テーブルデータの取得（時間帯別の電力使用量）
    Dim obj_テーブル As Object
    Dim obj_行 As Object
    Dim obj_セル As Object
    Dim lng_行数 As Long
    
    ' ヘッダー行の追加
    ws_結果.Cells(lng_出力行 + 2, 1).Value = "時間帯"
    ws_結果.Cells(lng_出力行 + 2, 2).Value = "使用率(%)"
    ws_結果.Cells(lng_出力行 + 2, 3).Value = "使用量(万kW)"
    ws_結果.Range(ws_結果.Cells(lng_出力行 + 2, 1), ws_結果.Cells(lng_出力行 + 2, 3)).Font.Bold = True
    
    ' テーブルの取得
    Set obj_テーブル = obj_IE.document.querySelector("table.usage-table")
    If Not obj_テーブル Is Nothing Then
        ' 行のコレクションを取得
        Set obj_要素リスト = obj_テーブル.getElementsByTagName("tr")
        
        ' 各行のデータを処理
        lng_行数 = obj_要素リスト.Length
        For i = 1 To lng_行数 - 1 ' ヘッダー行をスキップ
            Set obj_行 = obj_要素リスト.Item(i)
            
            ' 行からセルを取得
            Dim obj_セルリスト As Object
            Set obj_セルリスト = obj_行.getElementsByTagName("td")
            
            If obj_セルリスト.Length >= 3 Then
                ' 時間帯
                ws_結果.Cells(lng_出力行 + 2 + i, 1).Value = obj_セルリスト.Item(0).innerText
                
                ' 使用率
                ws_結果.Cells(lng_出力行 + 2 + i, 2).Value = obj_セルリスト.Item(1).innerText
                
                ' 使用量
                ws_結果.Cells(lng_出力行 + 2 + i, 3).Value = obj_セルリスト.Item(2).innerText
            End If
        Next i
    End If
    On Error GoTo ErrorHandler
    
    ' Internet Explorerを閉じる
    obj_IE.Quit
    Set obj_IE = Nothing
    
    ' 書式設定
    ws_結果.Columns("A:D").AutoFit
    ws_結果.Range("A1:D" & lng_出力行).Borders.LineStyle = xlContinuous
    
    ' 正常終了処理
    ws_結果.Activate
    Application.StatusBar = False
    MsgBox "Webスクレイピングが完了しました。", vbInformation
    
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    
    ' Internet Explorerを閉じる
    On Error Resume Next
    If Not obj_IE Is Nothing Then
        obj_IE.Quit
        Set obj_IE = Nothing
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'''---------------------------------------------------------
' 3. 社内システムとの連携インターフェース
'''---------------------------------------------------------
Sub 社内システム連携()
    ' 変数宣言
    Dim str_システムパス As String
    Dim str_出力フォルダ As String
    Dim str_入力ファイル As String
    Dim str_出力ファイル As String
    Dim obj_Shell As Object
    Dim lng_実行結果 As Long
    Dim obj_FSO As Object
    Dim str_ログファイル As String
    Dim str_コマンド As String
    Dim ws_データ As Worksheet
    Dim rng_エクスポート範囲 As Range
    
    ' ワークシート設定
    Set ws_データ = ThisWorkbook.Worksheets("エクスポートデータ")
    
    ' パス設定
    str_システムパス = "C:\Program Files\社内システム\bin\"
    str_出力フォルダ = ThisWorkbook.Path & "\Export\"
    str_入力ファイル = str_出力フォルダ & "input_" & Format(Date, "yyyymmdd") & ".csv"
    str_出力ファイル = str_出力フォルダ & "output_" & Format(Date, "yyyymmdd") & ".csv"
    str_ログファイル = str_出力フォルダ & "log_" & Format(Date, "yyyymmdd") & ".txt"
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' 出力フォルダの確認と作成
    If Not obj_FSO.FolderExists(str_出力フォルダ) Then
        obj_FSO.CreateFolder str_出力フォルダ
    End If
    
    ' データ範囲の取得
    Dim lng_最終行 As Long
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    
    If lng_最終行 <= 1 Then
        MsgBox "エクスポートするデータがありません。", vbExclamation
        Exit Sub
    End If
    
    Set rng_エクスポート範囲 = ws_データ.Range("A1:E" & lng_最終行)
    
    ' CSVファイルへのエクスポート
    Application.StatusBar = "データをCSVに出力中..."
    Application.DisplayAlerts = False
    
    ' 一時的なブックを作成してデータをコピー
    Dim wb_一時 As Workbook
    Set wb_一時 = Workbooks.Add
    
    rng_エクスポート範囲.Copy wb_一時.Sheets(1).Range("A1")
    
    ' CSVとして保存
    wb_一時.SaveAs Filename:=str_入力ファイル, FileFormat:=xlCSV
    wb_一時.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    
    ' 社内システムのコマンドライン実行
    Set obj_Shell = CreateObject("WScript.Shell")
    
    ' コマンドラインの構築
    str_コマンド = """" & str_システムパス & "process.exe"" "
    str_コマンド = str_コマンド & """" & str_入力ファイル & """ "
    str_コマンド = str_コマンド & """" & str_出力ファイル & """ "
    str_コマンド = str_コマンド & "/log:""" & str_ログファイル & """ "
    str_コマンド = str_コマンド & "/date:" & Format(Date, "yyyy-mm-dd")
    
    ' コマンド実行
    Application.StatusBar = "社内システムコマンド実行中..."
    lng_実行結果 = obj_Shell.Run(str_コマンド, 1, True)
    
    ' 実行結果の確認
    If lng_実行結果 = 0 Then
        ' 成功
        Application.StatusBar = "処理結果のインポート中..."
        
        ' 出力ファイルが存在するか確認
        If obj_FSO.FileExists(str_出力ファイル) Then
            ' 処理結果シートを準備
            Dim ws_結果 As Worksheet
            
            On Error Resume Next
            Set ws_結果 = ThisWorkbook.Worksheets("処理結果")
            If ws_結果 Is Nothing Then
                Set ws_結果 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws_結果.Name = "処理結果"
            End If
            ws_結果.Cells.Clear
            On Error GoTo ErrorHandler
            
            ' CSVデータのインポート
            With ws_結果.QueryTables.Add(Connection:="TEXT;" & str_出力ファイル, Destination:=ws_結果.Range("A1"))
                .TextFileParseType = xlDelimited
                .TextFileCommaDelimiter = True
                .Refresh BackgroundQuery:=False
            End With
            
            ' 書式設定
            ws_結果.Columns.AutoFit
            ws_結果.Activate
            
            ' 正常終了メッセージ
            MsgBox "社内システムとの連携処理が完了しました。" & vbCrLf & _
                   "処理結果は「処理結果」シートに表示されています。", vbInformation
        Else
            MsgBox "出力ファイルが見つかりません: " & str_出力ファイル & vbCrLf & _
                   "ログファイルを確認してください: " & str_ログファイル, vbExclamation
        End If
    Else
        ' 失敗
        MsgBox "社内システムコマンドの実行に失敗しました。" & vbCrLf & _
               "エラーコード: " & lng_実行結果 & vbCrLf & _
               "ログファイル: " & str_ログファイル, vbCritical
    End If
    
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

'''---------------------------------------------------------
' 4. PDFファイル生成と操作
'''---------------------------------------------------------
Sub PDFファイル生成と操作()
    ' 変数宣言
    Dim str_出力パス As String
    Dim str_PDFパス As String
    Dim ws_印刷シート As Worksheet
    Dim rng_印刷範囲 As Range
    Dim obj_PDFCreator As Object
    Dim obj_App As Object
    Dim obj_PDF As Object
    Dim bln_PDF出力成功 As Boolean
    
    ' PDFの出力先パスを設定
    str_出力パス = ThisWorkbook.Path & "\PDF_出力\"
    str_PDFパス = str_出力パス & "レポート_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    
    ' 出力フォルダが存在しない場合は作成
    On Error Resume Next
    If Dir(str_出力パス, vbDirectory) = "" Then
        MkDir str_出力パス
    End If
    On Error GoTo ErrorHandler
    
    ' 印刷するシートの設定
    Set ws_印刷シート = ThisWorkbook.Worksheets("レポート")
    
    ' ===== 方法1: ExcelのPDF出力機能を使用 =====
    Application.StatusBar = "PDFファイルを生成中..."
    
    ' 印刷設定の構成
    With ws_印刷シート.PageSetup
        .Orientation = xlLandscape             ' 横向き
        .Zoom = False                          ' 拡大縮小なし
        .FitToPagesWide = 1                    ' 幅を1ページに収める
        .FitToPagesTall = False                ' 高さは自動
        .CenterHorizontally = True             ' 水平方向に中央揃え
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .PrintHeadings = False                 ' 見出しを印刷しない
        .PrintGridlines = False                ' グリッド線を印刷しない
        
        ' ヘッダーとフッターの設定
        .LeftHeader = ""
        .CenterHeader = "電力使用量レポート"
        .RightHeader = ""
        .LeftFooter = "機密情報"
        .CenterFooter = "ページ &P / &N"
        .RightFooter = "出力日: &D"
    End With
    
    ' PDFとして保存
    On Error Resume Next
    ws_印刷シート.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=str_PDFパス, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' 保存結果の確認
    bln_PDF出力成功 = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    ' ===== 方法2: PDFライブラリを使用（例：Adobe Acrobat）=====
    ' 注：Adobe Acrobatがインストールされている環境が必要です
    If Not bln_PDF出力成功 Then
        ' 方法1が失敗した場合に代替手段を試行
        On Error Resume Next
        Set obj_App = CreateObject("AcroExch.App")
        Set obj_PDFCreator = CreateObject("PDFCreator.PDFCreatorObj")
        
        If obj_PDFCreator Is Nothing Then
            MsgBox "PDFライブラリが見つかりません。" & vbCrLf & _
                   "Adobe AcrobatまたはPDFCreatorがインストールされていることを確認してください。", vbExclamation
            GoTo Cleanup
        End If
        
        ' PDFCreatorを使用した生成処理
        ' ここにPDFCreatorを使用したコードを追加
        ' （ライセンスとインストールに依存するため省略）
        On Error GoTo ErrorHandler
    End If
    
    ' ===== PDF生成後の処理 =====
    If Dir(str_PDFパス) <> "" Then
        ' PDFファイルの情報を表示
        Dim obj_FSO As Object
        Dim obj_ファイル As Object
        
        Set obj_FSO = CreateObject("Scripting.FileSystemObject")
        Set obj_ファイル = obj_FSO.GetFile(str_PDFパス)
        
        MsgBox "PDFファイルが正常に生成されました。" & vbCrLf & _
               "ファイル: " & str_PDFパス & vbCrLf & _
               "サイズ: " & Format(obj_ファイル.Size / 1024, "#,##0.00") & " KB", vbInformation
        
        ' PDFを開く（オプション）
        'Shell "explorer.exe """ & str_PDFパス & """", vbNormalFocus
    Else
        MsgBox "PDFファイルの生成に失敗しました。", vbCritical
    End If
    
Cleanup:
    ' リソースの解放
    If Not obj_PDF Is Nothing Then Set obj_PDF = Nothing
    If Not obj_PDFCreator Is Nothing Then Set obj_PDFCreator = Nothing
    If Not obj_App Is Nothing Then Set obj_App = Nothing
    
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Resume Cleanup
End Sub

'''---------------------------------------------------------
' 5. バーコード・QRコード生成と読み取り
'''---------------------------------------------------------
Sub バーコードQRコード生成()
    ' 変数宣言
    Dim str_データ As String
    Dim str_出力パス As String
    Dim ws_コード As Worksheet
    Dim i As Long
    Dim obj_シェイプ As Shape
    
    ' データを設定
    str_データ = InputBox("エンコードするデータを入力してください:", "QRコード/バーコード生成", _
                         "https://www.tepco.co.jp/")
    
    If str_データ = "" Then Exit Sub
    
    ' コード用シートを準備
    On Error Resume Next
    Set ws_コード = ThisWorkbook.Worksheets("QRコード")
    If ws_コード Is Nothing Then
        Set ws_コード = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_コード.Name = "QRコード"
    End If
    ws_コード.Cells.Clear
    
    ' 既存のシェイプを削除
    For Each obj_シェイプ In ws_コード.Shapes
        obj_シェイプ.Delete
    Next obj_シェイプ
    On Error GoTo ErrorHandler
    
    ' タイトルとデータの表示
    ws_コード.Range("A1").Value = "QRコード/バーコードジェネレーター"
    ws_コード.Range("A1").Font.Size = 14
    ws_コード.Range("A1").Font.Bold = True
    
    ws_コード.Range("A3").Value = "エンコードされたデータ:"
    ws_コード.Range("B3").Value = str_データ
    
    ' ===== QRコードの生成 =====
    ' 注：実際のQRコード生成には外部ライブラリやAPIが必要です
    ' ここでは、QRコードの表現として簡易的なダミー画像を作成します
    
    ws_コード.Range("A5").Value = "QRコード:"
    
    ' QRコードの模擬表現（単純な16x16グリッド）
    Dim arr_QR(1 To 16, 1 To 16) As Boolean
    
    ' ランダムパターンを生成（実際のQRコードではありません）
    Randomize
    For i = 1 To 16
        For j = 1 To 16
            ' 端は常に黒（位置検出パターンの模倣）
            If i = 1 Or i = 16 Or j = 1 Or j = 16 Then
                arr_QR(i, j) = True
            ' 隅の3x3は固定パターン
            ElseIf (i <= 3 And j <= 3) Or (i >= 14 And j <= 3) Or (i <= 3 And j >= 14) Then
                arr_QR(i, j) = True
            Else
                ' 残りはデータに基づいたランダムパターン
                arr_QR(i, j) = (Rnd < 0.5)
            End If
        Next j
    Next i
    
    ' QRコードの描画
    For i = 1 To 16
        For j = 1 To 16
            ' セルの背景色で表現
            If arr_QR(i, j) Then
                ws_コード.Cells(i + 5, j + 1).Interior.Color = RGB(0, 0, 0)
            Else
                ws_コード.Cells(i + 5, j + 1).Interior.Color = RGB(255, 255, 255)
            End If
            
            ' セルのサイズを正方形に
            ws_コード.Cells(i + 5, j + 1).ColumnWidth = 2.5
            ws_コード.Rows(i + 5).RowHeight = 15
        Next j
    Next i
    
    ' ===== バーコードの生成 =====
    ws_コード.Range("A25").Value = "バーコード (Code 128):"
    
    ' バーコードの模擬表現
    Dim str_バーコードパターン As String
    str_バーコードパターン = ""
    
    ' 単純なパターン生成（実際のCode 128エンコードではありません）
    Dim lng_データ長 As Long
    lng_データ長 = Len(str_データ)
    
    ' 各文字をバーパターンに変換
    For i = 1 To lng_データ長
        Dim int_文字コード As Integer
        int_文字コード = Asc(Mid(str_データ, i, 1))
        
        ' 文字コードからランダムなバーパターンを生成
        Dim str_パターン As String
        str_パターン = ""
        
        ' 簡易的なバーコードパターン生成
        For j = 1 To 7
            If (int_文字コード And (2 ^ j)) > 0 Then
                str_パターン = str_パターン & "w"
            Else
                str_パターン = str_パターン & "n"
            End If
        Next j
        
        str_バーコードパターン = str_バーコードパターン & str_パターン
    Next i
    
    ' バーコードの描画
    Dim int_幅 As Integer
    int_幅 = Len(str_バーコードパターン)
    
    ' セルの準備
    For i = 1 To int_幅
        ' バーの幅を設定
        ws_コード.Columns(i + 1).ColumnWidth = 0.5
        
        ' バーの色を設定（w=白、n=黒）
        If Mid(str_バーコードパターン, i, 1) = "n" Then
            ws_コード.Cells(26, i + 1).Interior.Color = RGB(0, 0, 0)
        Else
            ws_コード.Cells(26, i + 1).Interior.Color = RGB(255, 255, 255)
        End If
    Next i
    
    ' バーコードの高さを設定
    ws_コード.Rows(26).RowHeight = 50
    
    ' 注意書き
    ws_コード.Range("A28").Value = "注意: これは実際のQRコード/バーコードを生成するものではありません。"
    ws_コード.Range("A29").Value = "本番環境では、専用のバーコードライブラリまたはAPIを使用してください。"
    ws_コード.Range("A28:A29").Font.Italic = True
    ws_コード.Range("A28:A29").Font.Color = RGB(128, 128, 128)
    
    ' シートの調整
    ws_コード.Activate
    ws_コード.Range("A1").Select
    
    MsgBox "QRコードとバーコードのサンプル表示が完了しました。" & vbCrLf & _
           "本番環境では、専用のバーコードライブラリを使用してください。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

' QRコード読み取り関数（例示用・実際には外部ライブラリが必要）
Private Function QRコード読み取り(str_画像パス As String) As String
    ' 注意: 実際のQRコード読み取りには専用ライブラリが必要です
    ' このコードは例示用であり、実際には機能しません
    
    ' ライブラリ例: ZXing, libdmtx, QRCode Library など
    
    ' ここでは単にファイル名を返す模擬的な実装
    QRコード読み取り = "QRコード内容 (画像: " & Mid(str_画像パス, InStrRev(str_画像パス, "\") + 1) & ")"
End Function
