Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim obj_最新ファイル As Object
    
    ' 初期化
    最新バックアップファイル検索 = ""
    
    ' FileSystemObjectの作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダが存在するか確認
    If Not obj_FSO.FolderExists(str_バックアップフォルダ) Then
        Exit Function
    End If
    
    ' フォルダオブジェクトの取得
    Set obj_フォルダ = obj_FSO.GetFolder(str_バックアップフォルダ)
    
    ' 最初のファイルを最新とする
    Set obj_最新ファイル = Nothing
    
    ' すべてのファイルをチェック
    For Each obj_ファイル In obj_フォルダ.Files
        ' バックアップファイルのみを対象に
        If InStr(1, obj_ファイル.Name, "バックアップ_", vbTextCompare) > 0 Then
            If obj_最新ファイル Is Nothing Then
                Set obj_最新ファイル = obj_ファイル
            ElseIf obj_ファイル.DateLastModified > obj_最新ファイル.DateLastModified Then
                Set obj_最新ファイル = obj_ファイル
            End If
        End If
    Next obj_ファイル
    
    ' 最新ファイルが見つかった場合はパスを返す
    If Not obj_最新ファイル Is Nothing Then
        最新バックアップファイル検索 = obj_最新ファイル.Path
    End If
End Function

'''---------------------------------------------------------
' 4. リソースリークを防ぐオブジェクト解放手法
'''---------------------------------------------------------
Sub リソースリーク防止()
    ' 変数宣言
    Dim ws_メモリ As Worksheet
    Dim obj_Excel As Object
    Dim obj_Word As Object
    Dim obj_Outlook As Object
    Dim col_プロセス As Object
    Dim obj_WMI As Object
    Dim obj_プロセス As Object
    Dim lng_Excel数 As Long
    Dim lng_Word数 As Long
    Dim lng_試行回数 As Long
    Dim arr_メモリ使用量() As Double
    
    ' メモリ監視用のワークシートを準備
    On Error Resume Next
    Set ws_メモリ = ThisWorkbook.Worksheets("メモリ監視")
    If ws_メモリ Is Nothing Then
        Set ws_メモリ = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_メモリ.Name = "メモリ監視"
    End If
    ws_メモリ.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' タイトル設定
    ws_メモリ.Range("A1").Value = "リソースリーク防止のデモ"
    ws_メモリ.Range("A1").Font.Size = 14
    ws_メモリ.Range("A1").Font.Bold = True
    
    ' ヘッダー行の設定
    ws_メモリ.Range("A3").Value = "プロセス情報"
    ws_メモリ.Range("A3").Font.Bold = True
    
    ws_メモリ.Range("A4").Value = "試行回数"
    ws_メモリ.Range("B4").Value = "Excel数"
    ws_メモリ.Range("C4").Value = "Word数"
    ws_メモリ.Range("D4").Value = "メモリ使用量"
    ws_メモリ.Range("A4:D4").Font.Bold = True
    
    ' 配列の初期化
    ReDim arr_メモリ使用量(1 To 10)
    
    ' ===== コンポーネント設計のベストプラクティス解説 =====
    ws_メモリ.Range("A15").Value = "オブジェクト解放のベストプラクティス:"
    ws_メモリ.Range("A15").Font.Bold = True
    
    ws_メモリ.Range("A16").Value = "1. Set obj = Nothing は必ず使用する"
    ws_メモリ.Range("A17").Value = "2. Exit Sub/Function の前に解放コードを配置"
    ws_メモリ.Range("A18").Value = "3. On Error GoTo によるクリーンアップを実装"
    ws_メモリ.Range("A19").Value = "4. 大きなオブジェクトはできるだけ早く解放"
    ws_メモリ.Range("A20").Value = "5. 入れ子のオブジェクトは内側から解放"
    
    ' ===== 実際のオブジェクト操作とリソース監視 =====
    For lng_試行回数 = 1 To 10
        ' ステータス表示
        Application.StatusBar = "リソース使用状況テスト実行中... " & lng_試行回数 & "/10"
        
        ' ===== 悪い例: リソースリークが発生するコード =====
        If lng_試行回数 <= 5 Then
            ws_メモリ.Range("A" & lng_試行回数 + 4).Value = lng_試行回数 & " (悪い例)"
            ws_メモリ.Range("A" & lng_試行回数 + 4).Interior.Color = RGB(255, 200, 200) ' 薄い赤色
            
            ' 意図的にオブジェクトを解放せずリークを発生させる
            Set obj_Word = CreateObject("Word.Application")
            ' ここで Set obj_Word = Nothing がない！
            
            ' このコメントは実際には実行しないでください（デモのみ）
            ' 実際のアプリケーションではこのようなコードはリソースリークの原因になります
        
        ' ===== 良い例: リソースリークを防止するコード =====
        Else
            ws_メモリ.Range("A" & lng_試行回数 + 4).Value = lng_試行回数 & " (良い例)"
            ws_メモリ.Range("A" & lng_試行回数 + 4).Interior.Color = RGB(200, 255, 200) ' 薄い緑色
            
            On Error Resume Next
            ' 正しいオブジェクト解放
            Set obj_Word = CreateObject("Word.Application")
            ' 使い終わったら必ず解放する
            If Not obj_Word Is Nothing Then
                obj_Word.Quit
                Set obj_Word = Nothing
            End If
            On Error GoTo ErrorHandler
        End If
        
        ' プロセス数とメモリ使用量の計測
        Call プロセス情報取得(lng_Excel数, lng_Word数, arr_メモリ使用量(lng_試行回数))
        
        ' 結果を表示
        ws_メモリ.Range("B" & lng_試行回数 + 4).Value = lng_Excel数
        ws_メモリ.Range("C" & lng_試行回数 + 4).Value = lng_Word数
        ws_メモリ.Range("D" & lng_試行回数 + 4).Value = arr_メモリ使用量(lng_試行回数)
        ws_メモリ.Range("D" & lng_試行回数 + 4).NumberFormat = "#,##0.00 MB"
        
        ' 少し待機
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next lng_試行回数
    
    ' ===== COMオブジェクトを適切に解放するコード例 =====
    ws_メモリ.Range("A30").Value = "正しいリソース解放パターン例:"
    ws_メモリ.Range("A30").Font.Bold = True
    
    ws_メモリ.Range("A31").Value = "' 良い例: 階層的なオブジェクト解放"
    ws_メモリ.Range("A32").Value = "If Not obj_OutlookItem Is Nothing Then Set obj_OutlookItem = Nothing"
    ws_メモリ.Range("A33").Value = "If Not obj_OutlookFolder Is Nothing Then Set obj_OutlookFolder = Nothing"
    ws_メモリ.Range("A34").Value = "If Not obj_Outlook Is Nothing Then"
    ws_メモリ.Range("A35").Value = "    obj_Outlook.Quit"
    ws_メモリ.Range("A36").Value = "    Set obj_Outlook = Nothing"
    ws_メモリ.Range("A37").Value = "End If"
    
    ws_メモリ.Range("A39").Value = "' 良い例: ファイルハンドルの解放"
    ws_メモリ.Range("A40").Value = "If iFileNum > 0 Then Close #iFileNum"
    
    ws_メモリ.Range("A42").Value = "' 良い例: エラー発生時も確実に解放するパターン"
    ws_メモリ.Range("A43").Value = "On Error GoTo ErrorHandler"
    ws_メモリ.Range("A44").Value = "' 処理コード..."
    ws_メモリ.Range("A45").Value = "CleanExit:"
    ws_メモリ.Range("A46").Value = "    ' リソース解放コード"
    ws_メモリ.Range("A47").Value = "    Exit Sub"
    ws_メモリ.Range("A48").Value = "ErrorHandler:"
    ws_メモリ.Range("A49").Value = "    ' エラー処理"
    ws_メモリ.Range("A50").Value = "    Resume CleanExit"
    
    ' 残っているオブジェクトを解放（重要：実際のアプリケーションでは常にこれを行う）
    On Error Resume Next
    If Not obj_Word Is Nothing Then
        obj_Word.Quit
        Set obj_Word = Nothing
    End If
    
    If Not obj_Excel Is Nothing Then
        Set obj_Excel = Nothing
    End If
    
    If Not obj_Outlook Is Nothing Then
        Set obj_Outlook = Nothing
    End If
    On Error GoTo ErrorHandler
    
    ' 書式設定
    ws_メモリ.Columns("A:D").AutoFit
    ws_メモリ.Range("A4:D14").Borders.LineStyle = xlContinuous
    
    ' グラフの作成
    Dim cht_メモリ使用量 As Chart
    Set cht_メモリ使用量 = Charts.Add
    
    With cht_メモリ使用量
        .ChartType = xlLine
        .SetSourceData Source:=ws_メモリ.Range("A4:A14,D4:D14")
        .HasTitle = True
        .ChartTitle.Text = "メモリ使用量の推移"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "メモリ使用量 (MB)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "試行回数"
        .Name = "メモリ使用量"
        .Location xlLocationAsObject, "メモリ監視"
    End With
    
    ' グラフの配置
    With ws_メモリ.ChartObjects(1)
        .Left = ws_メモリ.Range("F4").Left
        .Top = ws_メモリ.Range("F4").Top
        .Width = 400
        .Height = 250
    End With
    
    ' 結果を表示
    ws_メモリ.Activate
    ws_メモリ.Range("A1").Select
    Application.StatusBar = False
    
    MsgBox "リソースリーク防止のデモが完了しました。" & vbCrLf & _
           "「良い例」と「悪い例」のリソース管理の違いを確認してください。" & vbCrLf & _
           "実際のアプリケーションでは、常に適切なリソース解放を行ってください。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    
    ' リソース解放（エラー発生時も必ず実行）
    On Error Resume Next
    If Not obj_Word Is Nothing Then
        obj_Word.Quit
        Set obj_Word = Nothing
    End If
    
    If Not obj_Excel Is Nothing Then
        Set obj_Excel = Nothing
    End If
    
    If Not obj_Outlook Is Nothing Then
        Set obj_Outlook = Nothing
    End If
    
    Application.StatusBar = False
End Sub

' プロセス情報を取得する関数
Private Sub プロセス情報取得(ByRef lng_Excel数 As Long, ByRef lng_Word数 As Long, ByRef dbl_メモリ使用量 As Double)
    Dim obj_WMI As Object
    Dim col_プロセス As Object
    Dim obj_プロセス As Object
    Dim lng_合計メモリ As Long
    
    On Error Resume Next
    
    ' 初期化
    lng_Excel数 = 0
    lng_Word数 = 0
    dbl_メモリ使用量 = 0
    
    ' WMIサービスに接続
    Set obj_WMI = GetObject("winmgmts:\\.\root\cimv2")
    
    ' Excelプロセス数のカウント
    Set col_プロセス = obj_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    lng_Excel数 = col_プロセス.Count
    
    ' Wordプロセス数のカウント
    Set col_プロセス = obj_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'WINWORD.EXE'")
    lng_Word数 = col_プロセス.Count
    
    ' 現在のプロセスのメモリ使用量を取得
    Set col_プロセス = obj_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & GetCurrentProcessId())
    
    For Each obj_プロセス In col_プロセス
        ' WorkingSetSizeはバイト単位なのでMB単位に変換
        dbl_メモリ使用量 = obj_プロセス.WorkingSetSize / (1024 * 1024)
        Exit For
    Next obj_プロセス
    
    ' リソース解放
    Set col_プロセス = Nothing
    Set obj_WMI = Nothing
End Sub

' 現在のプロセスIDを取得するAPI関数
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

'''---------------------------------------------------------
' 5. バックアップとリカバリーの自動化
'''---------------------------------------------------------
Sub バックアップリカバリー自動化()
    ' 変数宣言
    Dim str_バックアップフォルダ As String
    Dim str_バックアップファイル As String
    Dim str_自動復元フォルダ As String
    Dim str_復元ファイル As String
    Dim obj_FSO As Object
    Dim bln_バックアップ成功 As Boolean
    Dim ws_ログ As Worksheet
    Dim i As Long
    
    ' ログ用のワークシートを準備
    On Error Resume Next
    Set ws_ログ = ThisWorkbook.Worksheets("バックアップログ")
    If ws_ログ Is Nothing Then
        Set ws_ログ = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_ログ.Name = "バックアップログ"
    End If
    ws_ログ.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' ===== パス設定 =====
    str_バックアップフォルダ = ThisWorkbook.Path & "\Backup\"
    str_自動復元フォルダ = ThisWorkbook.Path & "\Recovery\"
    
    ' FileSystemObjectの作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' バックアップフォルダの確認と作成
    If Not obj_FSO.FolderExists(str_バックアップフォルダ) Then
        obj_FSO.CreateFolder str_バックアップフォルダ
    End If
    
    ' 自動復元フォルダの確認と作成
    If Not obj_FSO.FolderExists(str_自動復元フォルダ) Then
        obj_FSO.CreateFolder str_自動復元フォルダ
    End If
    
    ' タイトル設定
    ws_ログ.Range("A1").Value = "バックアップ/リカバリー自動化ログ"
    ws_ログ.Range("A1").Font.Size = 14
    ws_ログ.Range("A1").Font.Bold = True
    
    ' ログのヘッダー行
    ws_ログ.Range("A3").Value = "日時"
    ws_ログ.Range("B3").Value = "操作"
    ws_ログ.Range("C3").Value = "ファイル"
    ws_ログ.Range("D3").Value = "結果"
    ws_ログ.Range("A3:D3").Font.Bold = True
    
    ' ===== 段階的バックアップの実装 =====
    
    ' 現在開いているブックのバックアップ
    Application.StatusBar = "バックアップを作成中..."
    
    ' バックアップ種別の判定
    Dim str_バックアップ種別 As String
    Dim lng_今日 As Long
    
    lng_今日 = CLng(Format(Date, "yyyymmdd"))
    
    ' 曜日に基づいて異なるバックアップ種別を判定
    Select Case Weekday(Date)
        Case vbSunday
            str_バックアップ種別 = "週次"
        Case vbMonday
            str_バックアップ種別 = "増分"
        Case vbFriday
            str_バックアップ種別 = "金曜日"
        Case Else
            str_バックアップ種別 = "日次"
    End Select
    
    ' バックアップファイルパスの作成
    str_バックアップファイル = str_バックアップフォルダ & str_バックアップ種別 & "_" & _
                             ThisWorkbook.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    
    ' 一時的に表示警告をオフ
    Application.DisplayAlerts = False
    
    ' バックアップを保存
    ThisWorkbook.SaveCopyAs str_バックアップファイル
    
    ' バックアップ結果をログに記録
    ws_ログ.Range("A4").Value = Now
    ws_ログ.Range("B4").Value = "バックアップ作成"
    ws_ログ.Range("C4").Value = str_バックアップファイル
    ws_ログ.Range("D4").Value = "成功"
    ws_ログ.Range("A4:D4").Interior.Color = RGB(200, 255, 200)  ' 薄い緑色
    
    ' 表示警告を元に戻す
    Application.DisplayAlerts = True
    
    ' ===== 古いバックアップの整理 =====
    Call 古いバックアップファイル整理(str_バックアップフォルダ, ws_ログ, 5)
    
    ' ===== バックアップからのリカバリーのデモ =====
    Application.StatusBar = "リカバリーのデモを実行中..."
    
    ' リカバリーのデモを実行
    bln_バックアップ成功 = False
    
    ' 最新のバックアップファイルを取得
    str_復元ファイル = 最新バックアップファイル検索(str_バックアップフォルダ)
    
    If str_復元ファイル <> "" Then
        ' バックアップファイルが存在する場合
        bln_バックアップ成功 = True
        
        ' リカバリー用ディレクトリにコピー
        Dim str_復元先 As String
        str_復元先 = str_自動復元フォルダ & "復元_" & Format(Now, "yyyymmdd_hhnnss") & "_" & _
                    obj_FSO.GetFileName(str_復元ファイル)
        
        obj_FSO.CopyFile str_復元ファイル, str_復元先
        
        ' リカバリーログを記録
        ws_ログ.Range("A5").Value = Now
        ws_ログ.Range("B5").Value = "リカバリーデモ"
        ws_ログ.Range("C5").Value = str_復元先
        ws_ログ.Range("D5").Value = "成功"
        ws_ログ.Range("A5:D5").Interior.Color = RGB(200, 255, 200)  ' 薄い緑色
    Else
        ' バックアップファイルが存在しない場合
        ws_ログ.Range("A5").Value = Now
        ws_ログ.Range("B5").Value = "リカバリーデモ"
        ws_ログ.Range("C5").Value = "該当なし"
        ws_ログ.Range("D5").Value = "失敗 - バックアップが見つかりません"
        ws_ログ.Range("A5:D5").Interior.Color = RGB(255, 200, 200)  ' 薄い赤色
    End If
    
    ' ===== バックアップスケジュールの解説 =====
    ws_ログ.Range("A8").Value = "推奨バックアップスケジュール:"
    ws_ログ.Range("A8").Font.Bold = True
    
    ws_ログ.Range("A9").Value = "1. 日次バックアップ: 毎日作業終了時に実行"
    ws_ログ.Range("A10").Value = "2. 週次バックアップ: 毎週金曜日に実行（前週分は保持）"
    ws_ログ.Range("A11").Value = "3. 月次バックアップ: 月末に実行（前月分は3ヵ月保持）"
    ws_ログ.Range("A12").Value = "4. アーカイブバックアップ: 四半期ごとに実行（永久保存）"
    
    ' ===== 自動バックアップの設定方法 =====
    ws_ログ.Range("A14").Value = "自動バックアップの設定方法:"
    ws_ログ.Range("A14").Font.Bold = True
    
    ws_ログ.Range("A15").Value = "1. Workbook_BeforeCloseイベントでバックアップを実行"
    ws_ログ.Range("A16").Value = "2. Windows タスクスケジューラを使用"
    ws_ログ.Range("A17").Value = "3. バックアップ状態を監視するログ機能の実装"
    ws_ログ.Range("A18").Value = "4. リカバリー手順の文書化と定期的なテスト"
    
    ' Workbook_BeforeCloseの例
    ws_ログ.Range("A20").Value = "' ThisWorkbookモジュールの例:"
    ws_ログ.Range("A21").Value = "Private Sub Workbook_BeforeClose(Cancel As Boolean)"
    ws_ログ.Range("A22").Value = "    ' ブック閉じる前に自動バックアップを実行"
    ws_ログ.Range("A23").Value = "    Call バックアップリカバリー自動化"
    ws_ログ.Range("A24").Value = "End Sub"
    
    ' 書式設定
    ws_ログ.Columns("A:D").AutoFit
    ws_ログ.Range("A3:D5").Borders.LineStyle = xlContinuous
    
    ' 結果の表示
    ws_ログ.Activate
    ws_ログ.Range("A1").Select
    
    Application.StatusBar = False
    
    ' 完了メッセージ
    MsgBox "バックアップ処理が完了しました。" & vbCrLf & _
           IIf(bln_バックアップ成功, "リカバリーデモも正常に実行されました。", "リカバリーデモは失敗しました。有効なバックアップが見つかりません。") & vbCrLf & vbCrLf & _
           "詳細はバックアップログシートを確認してください。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.DisplayAlerts = True
    Application.StatusBar = False
End Sub

' 古いバックアップファイルを整理する関数
Private Sub 古いバックアップファイル整理(str_フォルダパス As String, ws_ログ As Worksheet, lng_保持日数 As Long)
    Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim dt_基準日 As Date
    Dim lng_削除数 As Long
    Dim lng_ログ行 As Long
    
    ' 初期化
    lng_削除数 = 0
    lng_ログ行 = 6
    dt_基準日 = DateAdd("d", -lng_保持日数, Date)
    
    ' FileSystemObjectの作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダが存在するか確認
    If Not obj_FSO.FolderExists(str_フォルダパス) Then
        Exit Sub
    End If
    
    ' フォルダオブジェクトの取得
    Set obj_フォルダ = obj_FSO.GetFolder(str_フォルダパス)
    
    ' すべてのファイルをチェック
    For Each obj_ファイル In obj_フォルダ.Files
        ' バックアップファイルのみを対象に
        If InStr(1, obj_ファイル.Name, "日次_", vbTextCompare) > 0 Then
            ' 基準日より古いファイルを削除
            If obj_ファイル.DateCreated < dt_基準日 Then
                ' バックアップ整理をログに記録
                ws_ログ.Range("A" & lng_ログ行).Value = Now
                ws_ログ.Range("B" & lng_ログ行).Value = "古いバックアップ削除"
                ws_ログ.Range("C" & lng_ログ行).Value = obj_ファイル.Name
                
                On Error Resume Next
                obj_FSO.DeleteFile obj_ファイル.Path, True
                
                If Err.Number = 0 Then
                    ws_ログ.Range("D" & lng_ログ行).Value = "成功"
                    ws_ログ.Range("A" & lng_ログ行 & ":D" & lng_ログ行).Interior.Color = RGB(240, 240, 240)  ' 薄いグレー
                    lng_削除数 = lng_削除数 + 1
                Else
                    ws_ログ.Range("D" & lng_ログ行).Value = "失敗 - " & Err.Description
                    ws_ログ.Range("A" & lng_ログ行 & ":D" & lng_ログ行).Interior.Color = RGB(255, 200, 200)  ' 薄い赤色
                End If
                On Error GoTo 0
                
                lng_ログ行 = lng_ログ行 + 1
            End If
        End If
    Next obj_ファイル
    
    ' バックアップ整理の結果をログに記録
    ws_ログ.Range("A" & lng_ログ行).Value = Now
    ws_ログ.Range("B" & lng_ログ行).Value = "バックアップ整理サマリー"
    ws_ログ.Range("C" & lng_ログ行).Value = str_フォルダパス
    ws_ログ.Range("D" & lng_ログ行).Value = lng_削除数 & " 件のファイルを削除しました。"
    ws_ログ.Range("A" & lng_ログ行 & ":D" & lng_ログ行).Interior.Color = RGB(220, 220, 240)  ' 薄い青色
End Sub    If ws_管理 Is Nothing Then
        Set ws_管理 = ThisWorkbook.Worksheets.Add(After:=ws_フォーム)
        ws_管理.Name = "管理"
        ws_管理.Visible = xlSheetVeryHidden  ' 管理シートは非表示に
    End If
    On Error GoTo ErrorHandler
    
    ' ===== 入力フォーム用のシートを設定 =====
    ws_フォーム.Cells.Clear
    
    ' タイトル設定
    ws_フォーム.Range("A1").Value = "供給地点登録フォーム"
    ws_フォーム.Range("A1").Font.Size = 14
    ws_フォーム.Range("A1").Font.Bold = True
    
    ' フォームのヘッダー
    ws_フォーム.Range("A3").Value = "供給地点特定番号:"
    ws_フォーム.Range("A4").Value = "顧客名:"
    ws_フォーム.Range("A5").Value = "電力会社:"
    ws_フォーム.Range("A6").Value = "契約種別:"
    ws_フォーム.Range("A7").Value = "契約容量(kW):"
    ws_フォーム.Range("A8").Value = "供給開始日:"
    
    ' 入力欄の設定
    Set rng_入力範囲 = ws_フォーム.Range("B3:B8")
    
    ' ===== 入力規則と入力支援の設定 =====
    
    ' 供給地点特定番号の入力規則（22桁の数字）
    With ws_フォーム.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateCustom, Formula1:="=AND(LEN(B3)=22,ISNUMBER(--SUBSTITUTE(B3,""-"","""")))"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "供給地点特定番号は22桁の数字（ハイフンなし）で入力してください。"
        .InputTitle = "供給地点特定番号"
        .InputMessage = "22桁の数字（ハイフンなし）で入力してください。" & vbCrLf & "例：1234567890123456789012"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 自動書式設定（数字22桁を見やすく区切る）
    ws_フォーム.Range("B3").NumberFormat = "00\-0000\-0000\-0000\-0000\-00"
    
    ' 顧客名（必須入力）
    With ws_フォーム.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateInputRequired
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "顧客名は必須項目です。入力してください。"
        .InputTitle = "顧客名"
        .InputMessage = "顧客名を入力してください。"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 電力会社（ドロップダウンリスト）
    arr_電力会社 = Array("東京電力", "関西電力", "中部電力", "東北電力", "北海道電力", "北陸電力", "中国電力", "四国電力", "九州電力", "沖縄電力")
    
    ' 管理シートにリストを設定
    ws_管理.Range("A1").Value = "電力会社リスト"
    For i = 0 To UBound(arr_電力会社)
        ws_管理.Cells(i + 2, 1).Value = arr_電力会社(i)
    Next i
    
    ' 名前付き範囲を定義
    ThisWorkbook.Names.Add Name:="電力会社リスト", RefersTo:=ws_管理.Range("A2:A" & (UBound(arr_電力会社) + 2))
    
    ' ドロップダウンリストを設定
    With ws_フォーム.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=電力会社リスト"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "リストから電力会社を選択してください。"
        .InputTitle = "電力会社"
        .InputMessage = "リストから選択してください。"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 契約種別（ドロップダウンリスト）
    arr_契約種別 = Array("従量電灯A", "従量電灯B", "従量電灯C", "低圧電力", "高圧電力A", "高圧電力B", "特別高圧電力")
    
    ' 管理シートにリストを設定
    ws_管理.Range("B1").Value = "契約種別リスト"
    For i = 0 To UBound(arr_契約種別)
        ws_管理.Cells(i + 2, 2).Value = arr_契約種別(i)
    Next i
    
    ' 名前付き範囲を定義
    ThisWorkbook.Names.Add Name:="契約種別リスト", RefersTo:=ws_管理.Range("B2:B" & (UBound(arr_契約種別) + 2))
    
    ' ドロップダウンリストを設定
    With ws_フォーム.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=契約種別リスト"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "リストから契約種別を選択してください。"
        .InputTitle = "契約種別"
        .InputMessage = "リストから選択してください。"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 契約容量（数値制限）
    With ws_フォーム.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateDecimal, Operator:=xlGreater, Formula1:="0"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "契約容量は0より大きい数値を入力してください。"
        .InputTitle = "契約容量(kW)"
        .InputMessage = "正の数値を入力してください。"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 供給開始日（日付制限）
    With ws_フォーム.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateDate, Operator:=xlGreaterEqual, Formula1:="=TODAY()-365"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "有効な日付を入力してください（過去1年以内〜未来）。"
        .InputTitle = "供給開始日"
        .InputMessage = "日付を入力してください。" & vbCrLf & "例: " & Format(Date, "yyyy/mm/dd")
        .ShowInput = True
        .ShowError = True
    End With
    
    ' 日付形式の設定
    ws_フォーム.Range("B8").NumberFormat = "yyyy/mm/dd"
    
    ' ===== データ登録用ボタンの追加 =====
    Dim btn_登録 As Button
    
    ' 既存のボタンを削除
    On Error Resume Next
    For Each btn In ws_フォーム.Buttons
        btn.Delete
    Next btn
    On Error GoTo ErrorHandler
    
    ' 登録ボタンを追加
    Set btn_登録 = ws_フォーム.Buttons.Add(200, 200, 80, 30)
    With btn_登録
        .Caption = "登録"
        .Name = "btnSubmit"
        .OnAction = "データ登録実行"
    End With
    
    ' ボタンの位置調整
    btn_登録.Left = ws_フォーム.Range("B10").Left
    btn_登録.Top = ws_フォーム.Range("B10").Top
    
    ' 書式設定
    ws_フォーム.Columns("A:B").AutoFit
    ws_フォーム.Range("A3:B8").Borders.LineStyle = xlContinuous
    
    ' 入力支援機能についての説明
    str_メッセージ = "入力フォームに以下の入力支援機能を設定しました：" & vbCrLf & vbCrLf & _
                    "1. 供給地点特定番号: 22桁の数字チェックと自動書式設定" & vbCrLf & _
                    "2. 顧客名: 必須入力チェック" & vbCrLf & _
                    "3. 電力会社: ドロップダウンリストから選択" & vbCrLf & _
                    "4. 契約種別: ドロップダウンリストから選択" & vbCrLf & _
                    "5. 契約容量: 正の数値チェック" & vbCrLf & _
                    "6. 供給開始日: 日付形式チェックと範囲制限" & vbCrLf & vbCrLf & _
                    "各フィールドにマウスを重ねると入力ガイドが表示されます。"
    
    MsgBox str_メッセージ, vbInformation, "入力フォーム設定完了"
    
    ws_フォーム.Activate
    ws_フォーム.Range("B3").Select
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

' データ登録実行サブルーチン（ボタンから呼び出される）
Sub データ登録実行()
    ' 変数宣言
    Dim ws_フォーム As Worksheet
    Dim ws_データ As Worksheet
    Dim str_供給地点番号 As String
    Dim str_顧客名 As String
    Dim str_電力会社 As String
    Dim str_契約種別 As String
    Dim dbl_契約容量 As Double
    Dim dt_供給開始日 As Date
    Dim lng_次行 As Long
    Dim bln_エラーあり As Boolean
    Dim str_エラーメッセージ As String
    
    ' ワークシート設定
    On Error Resume Next
    Set ws_フォーム = ThisWorkbook.Worksheets("入力フォーム")
    Set ws_データ = ThisWorkbook.Worksheets("登録データ")
    
    If ws_データ Is Nothing Then
        Set ws_データ = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_データ.Name = "登録データ"
        
        ' ヘッダー行の設定
        ws_データ.Range("A1").Value = "供給地点特定番号"
        ws_データ.Range("B1").Value = "顧客名"
        ws_データ.Range("C1").Value = "電力会社"
        ws_データ.Range("D1").Value = "契約種別"
        ws_データ.Range("E1").Value = "契約容量(kW)"
        ws_データ.Range("F1").Value = "供給開始日"
        ws_データ.Range("G1").Value = "登録日時"
        ws_データ.Range("H1").Value = "登録ユーザー"
        
        ws_データ.Range("A1:H1").Font.Bold = True
    End If
    On Error GoTo ErrorHandler
    
    ' 入力値の取得とバリデーション
    str_供給地点番号 = ws_フォーム.Range("B3").Value
    str_顧客名 = ws_フォーム.Range("B4").Value
    str_電力会社 = ws_フォーム.Range("B5").Value
    str_契約種別 = ws_フォーム.Range("B6").Value
    
    ' 数値・日付の変換と検証
    bln_エラーあり = False
    str_エラーメッセージ = "以下の項目を確認してください：" & vbCrLf
    
    ' 供給地点番号のチェック（22桁の数字）
    If Len(Replace(str_供給地点番号, "-", "")) <> 22 Or Not IsNumeric(Replace(str_供給地点番号, "-", "")) Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 供給地点特定番号は22桁の数字で入力してください" & vbCrLf
    End If
    
    ' 顧客名のチェック（必須）
    If Trim(str_顧客名) = "" Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 顧客名は必須です" & vbCrLf
    End If
    
    ' 電力会社のチェック（リストから選択）
    If Trim(str_電力会社) = "" Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 電力会社を選択してください" & vbCrLf
    End If
    
    ' 契約種別のチェック（リストから選択）
    If Trim(str_契約種別) = "" Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 契約種別を選択してください" & vbCrLf
    End If
    
    ' 契約容量のチェック（正の数値）
    On Error Resume Next
    dbl_契約容量 = CDbl(ws_フォーム.Range("B7").Value)
    If Err.Number <> 0 Or dbl_契約容量 <= 0 Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 契約容量は正の数値で入力してください" & vbCrLf
    End If
    Err.Clear
    
    ' 供給開始日のチェック（有効な日付）
    On Error Resume Next
    dt_供給開始日 = CDate(ws_フォーム.Range("B8").Value)
    If Err.Number <> 0 Then
        bln_エラーあり = True
        str_エラーメッセージ = str_エラーメッセージ & "- 有効な供給開始日を入力してください" & vbCrLf
    End If
    On Error GoTo ErrorHandler
    
    ' エラーがある場合は処理を中止
    If bln_エラーあり Then
        MsgBox str_エラーメッセージ, vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' 次の登録行を取得
    lng_次行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row + 1
    
    ' データを登録
    ws_データ.Cells(lng_次行, 1).Value = str_供給地点番号
    ws_データ.Cells(lng_次行, 2).Value = str_顧客名
    ws_データ.Cells(lng_次行, 3).Value = str_電力会社
    ws_データ.Cells(lng_次行, 4).Value = str_契約種別
    ws_データ.Cells(lng_次行, 5).Value = dbl_契約容量
    ws_データ.Cells(lng_次行, 6).Value = dt_供給開始日
    ws_データ.Cells(lng_次行, 7).Value = Now
    ws_データ.Cells(lng_次行, 8).Value = Application.UserName
    
    ' 書式設定
    ws_データ.Cells(lng_次行, 1).NumberFormat = "00\-0000\-0000\-0000\-0000\-00"
    ws_データ.Cells(lng_次行, 5).NumberFormat = "#,##0.00"
    ws_データ.Cells(lng_次行, 6).NumberFormat = "yyyy/mm/dd"
    ws_データ.Cells(lng_次行, 7).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    
    ' データ登録後のフォームクリア
    ws_フォーム.Range("B3:B8").ClearContents
    
    ' 成功メッセージ
    MsgBox "データが正常に登録されました。", vbInformation, "登録完了"
    
    ' フォームの先頭フィールドにフォーカス
    ws_フォーム.Activate
    ws_フォーム.Range("B3").Select
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "データ登録中にエラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

'''---------------------------------------------------------
' 3. ネットワーク切断・再接続時の復旧処理
'''---------------------------------------------------------
Sub ネットワーク切断復旧処理()
    ' 変数宣言
    Dim str_サーバーパス As String
    Dim str_ファイルパス As String
    Dim str_バックアップパス As String
    Dim int_リトライ回数 As Integer
    Dim int_最大リトライ回数 As Integer
    Dim int_待機時間_秒 As Integer
    Dim bln_接続成功 As Boolean
    Dim lngTimer As Long
    Dim obj_FSO As Object
    Dim ws_状態 As Worksheet
    
    ' 状態表示用のワークシートを準備
    On Error Resume Next
    Set ws_状態 = ThisWorkbook.Worksheets("接続状態")
    If ws_状態 Is Nothing Then
        Set ws_状態 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_状態.Name = "接続状態"
    End If
    ws_状態.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' ===== パスとリトライ設定 =====
    str_サーバーパス = "\\server\share\"  ' 実際のサーバーパスに変更
    str_ファイルパス = str_サーバーパス & "データ\マスタ.xlsx"
    str_バックアップパス = ThisWorkbook.Path & "\バックアップ\"
    
    ' リトライ設定
    int_最大リトライ回数 = 5      ' 最大リトライ回数
    int_待機時間_秒 = 3           ' リトライ間の待機時間（秒）
    
    ' ステータス表示の設定
    ws_状態.Range("A1").Value = "ネットワーク接続状態モニター"
    ws_状態.Range("A1").Font.Size = 14
    ws_状態.Range("A1").Font.Bold = True
    
    ws_状態.Range("A3").Value = "サーバーパス:"
    ws_状態.Range("B3").Value = str_サーバーパス
    
    ws_状態.Range("A4").Value = "ターゲットファイル:"
    ws_状態.Range("B4").Value = str_ファイルパス
    
    ws_状態.Range("A5").Value = "バックアップパス:"
    ws_状態.Range("B5").Value = str_バックアップパス
    
    ws_状態.Range("A7").Value = "接続ステータス:"
    ws_状態.Range("B7").Value = "確認中..."
    
    ws_状態.Range("A8").Value = "リトライ回数:"
    ws_状態.Range("B8").Value = "0 / " & int_最大リトライ回数
    
    ws_状態.Range("A9").Value = "最終確認時刻:"
    ws_状態.Range("B9").Value = Now
    ws_状態.Range("B9").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    
    ' FileSystemObjectの作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' バックアップフォルダの確認と作成
    If Not obj_FSO.FolderExists(str_バックアップパス) Then
        obj_FSO.CreateFolder str_バックアップパス
    End If
    
    ' ===== ネットワーク接続チェックと再接続処理 =====
    Application.StatusBar = "ネットワーク接続を確認中..."
    
    ' 接続チェックの初期化
    bln_接続成功 = False
    int_リトライ回数 = 0
    
    ' 接続が成功するか最大リトライ回数に達するまで繰り返す
    Do While Not bln_接続成功 And int_リトライ回数 <= int_最大リトライ回数
        ' リトライ回数の更新
        int_リトライ回数 = int_リトライ回数 + 1
        ws_状態.Range("B8").Value = int_リトライ回数 & " / " & int_最大リトライ回数
        
        ' ステータス更新
        ws_状態.Range("B7").Value = "接続試行中... (" & int_リトライ回数 & "回目)"
        ws_状態.Range("B9").Value = Now
        
        ' ネットワークパスの存在チェック
        On Error Resume Next
        bln_接続成功 = obj_FSO.FolderExists(str_サーバーパス)
        
        If bln_接続成功 Then
            ' 接続成功
            ws_状態.Range("B7").Value = "接続成功"
            ws_状態.Range("B7").Interior.Color = RGB(200, 255, 200) ' 薄い緑色
            
            ' ネットワーク上のファイルにアクセス
            If obj_FSO.FileExists(str_ファイルパス) Then
                ' ファイルが存在する場合の処理
                ws_状態.Range("A11").Value = "ファイルステータス:"
                ws_状態.Range("B11").Value = "ファイルにアクセス可能"
                ws_状態.Range("B11").Interior.Color = RGB(200, 255, 200) ' 薄い緑色
                
                ' ファイル情報の取得
                Dim obj_ファイル As Object
                Set obj_ファイル = obj_FSO.GetFile(str_ファイルパス)
                
                ws_状態.Range("A12").Value = "ファイルサイズ:"
                ws_状態.Range("B12").Value = Format(obj_ファイル.Size / 1024, "#,##0.00") & " KB"
                
                ws_状態.Range("A13").Value = "最終更新日時:"
                ws_状態.Range("B13").Value = obj_ファイル.DateLastModified
                ws_状態.Range("B13").NumberFormat = "yyyy/mm/dd hh:mm:ss"
                
                ' バックアップ作成（オプション）
                Dim str_バックアップファイル As String
                str_バックアップファイル = str_バックアップパス & "バックアップ_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
                
                obj_FSO.CopyFile str_ファイルパス, str_バックアップファイル
                
                ws_状態.Range("A15").Value = "バックアップステータス:"
                ws_状態.Range("B15").Value = "バックアップ作成成功"
                ws_状態.Range("B15").Interior.Color = RGB(200, 255, 200) ' 薄い緑色
                
                ws_状態.Range("A16").Value = "バックアップファイル:"
                ws_状態.Range("B16").Value = str_バックアップファイル
            Else
                ' ファイルが存在しない場合の処理
                ws_状態.Range("A11").Value = "ファイルステータス:"
                ws_状態.Range("B11").Value = "ファイルが見つかりません"
                ws_状態.Range("B11").Interior.Color = RGB(255, 200, 200) ' 薄い赤色
            End If
        Else
            ' 接続失敗
            ws_状態.Range("B7").Value = "接続失敗 - リトライ中..."
            ws_状態.Range("B7").Interior.Color = RGB(255, 200, 200) ' 薄い赤色
            
            ' エラー情報をクリア
            Err.Clear
            
            ' 次のリトライの前に待機
            If int_リトライ回数 < int_最大リトライ回数 Then
                Application.StatusBar = "接続再試行まで " & int_待機時間_秒 & " 秒待機中..."
                
                ' 待機時間のカウントダウン表示
                lngTimer = Timer
                Do While Timer < lngTimer + int_待機時間_秒
                    DoEvents
                    ws_状態.Range("A10").Value = "再試行まで:"
                    ws_状態.Range("B10").Value = Format(lngTimer + int_待機時間_秒 - Timer, "0") & " 秒"
                Loop
            End If
        End If
    Loop
    On Error GoTo ErrorHandler
    
    ' 最終結果の処理
    If bln_接続成功 Then
        MsgBox "ネットワーク接続に成功しました。" & vbCrLf & _
               "リトライ回数: " & int_リトライ回数 & " / " & int_最大リトライ回数, vbInformation
    Else
        ' すべてのリトライが失敗した場合
        ws_状態.Range("B7").Value = "接続失敗 - リトライ上限到達"
        ws_状態.Range("B7").Interior.Color = RGB(255, 150, 150) ' 濃い赤色
        
        ' オフラインモードでの代替処理（例：ローカルキャッシュの使用）
        ws_状態.Range("A18").Value = "オフラインモード:"
        ws_状態.Range("B18").Value = "有効"
        
        ' 最新のバックアップファイルを検索
        Dim str_最新バックアップ As String
        str_最新バックアップ = 最新バックアップファイル検索(str_バックアップパス)
        
        If str_最新バックアップ <> "" Then
            ws_状態.Range("A19").Value = "使用可能なバックアップ:"
            ws_状態.Range("B19").Value = str_最新バックアップ
            
            Dim obj_バックアップファイル As Object
            Set obj_バックアップファイル = obj_FSO.GetFile(str_最新バックアップ)
            
            ws_状態.Range("A20").Value = "バックアップ日時:"
            ws_状態.Range("B20").Value = obj_バックアップファイル.DateLastModified
            ws_状態.Range("B20").NumberFormat = "yyyy/mm/dd hh:mm:ss"
        Else
            ws_状態.Range("A19").Value = "使用可能なバックアップ:"
            ws_状態.Range("B19").Value = "なし"
        End If
        
        MsgBox "ネットワーク接続に失敗しました。オフラインモードで動作します。" & vbCrLf & _
               "ネットワーク管理者に連絡してください。", vbExclamation
    End If
    
    ' 書式設定
    ws_状態.Columns("A:B").AutoFit
    ws_状態.Activate
    ws_状態.Range("A1").Select
    
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

' 最新のバックアップファイルを検索する関数
Private Function 最新バックアップファイル検索(str_バックアップフォルダ As String) As String
    Dim obj_FSO As Object
    Dim obj_フォルダ As Object
    Dim obj_ファイル As Object
    Dim obj_最新ファイル As Object
    '''---------------------------------------------------------
' 1. 実行環境の違いに対応するコード設計
'''---------------------------------------------------------
Sub 実行環境対応設計()
    ' 変数宣言
    Dim str_実行環境情報 As String
    Dim lng_Excel行数上限 As Long
    Dim lng_Excel列数上限 As Long
    Dim str_OSバージョン As String
    Dim str_Office言語 As String
    Dim bln_64ビット環境 As Boolean
    Dim str_使用可能API As String
    Dim bln_MAC環境 As Boolean
    Dim ws_情報 As Worksheet
    
    ' 情報表示用のワークシートを準備
    On Error Resume Next
    Set ws_情報 = ThisWorkbook.Worksheets("環境情報")
    If ws_情報 Is Nothing Then
        Set ws_情報 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_情報.Name = "環境情報"
    End If
    ws_情報.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' タイトルの設定
    ws_情報.Range("A1").Value = "実行環境情報"
    ws_情報.Range("A1").Font.Size = 14
    ws_情報.Range("A1").Font.Bold = True
    
    ' ===== 環境情報の収集 =====
    
    ' Excelバージョンの取得
    ws_情報.Range("A3").Value = "Excelバージョン:"
    ws_情報.Range("B3").Value = Application.Version
    
    ' Excelのビルド番号
    ws_情報.Range("A4").Value = "Excelビルド:"
    ws_情報.Range("B4").Value = Application.Build
    
    ' 32bit/64bit環境の判定
    ws_情報.Range("A5").Value = "アーキテクチャ:"
    
    #If Win64 Then
        bln_64ビット環境 = True
        ws_情報.Range("B5").Value = "64ビット"
    #Else
        bln_64ビット環境 = False
        ws_情報.Range("B5").Value = "32ビット"
    #End If
    
    ' Windows/MAC環境の判定
    ws_情報.Range("A6").Value = "OS環境:"
    
    #If Mac Then
        bln_MAC環境 = True
        ws_情報.Range("B6").Value = "macOS"
    #Else
        bln_MAC環境 = False
        ws_情報.Range("B6").Value = "Windows"
    #End If
    
    ' OSバージョンの取得（Windows環境のみ）
    ws_情報.Range("A7").Value = "OSバージョン:"
    
    If Not bln_MAC環境 Then
        On Error Resume Next
        ' WMIを使用してOSバージョンを取得
        Dim obj_WMIService As Object
        Dim col_OS As Object
        Dim obj_OS As Object
        
        Set obj_WMIService = GetObject("winmgmts:\\.\root\cimv2")
        Set col_OS = obj_WMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each obj_OS In col_OS
            str_OSバージョン = obj_OS.Caption & " (Build " & obj_OS.BuildNumber & ")"
            Exit For
        Next
        
        If Err.Number <> 0 Then
            str_OSバージョン = "取得できません"
        End If
        On Error GoTo ErrorHandler
    Else
        str_OSバージョン = "macOS (バージョン不明)"
    End If
    
    ws_情報.Range("B7").Value = str_OSバージョン
    
    ' Office言語設定の取得
    ws_情報.Range("A8").Value = "Office言語:"
    str_Office言語 = Switch( _
        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1033, "英語 (米国)", _
        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1041, "日本語", _
        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1031, "ドイツ語", _
        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1036, "フランス語", _
        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1034, "スペイン語", _
        True, "その他 (ID: " & Application.LanguageSettings.LanguageID(msoLanguageIDUI) & ")")
    
    ws_情報.Range("B8").Value = str_Office言語
    
    ' Excelの制限値を取得
    ws_情報.Range("A10").Value = "行数上限:"
    lng_Excel行数上限 = Application.Rows.Count
    ws_情報.Range("B10").Value = lng_Excel行数上限
    
    ws_情報.Range("A11").Value = "列数上限:"
    lng_Excel列数上限 = Application.Columns.Count
    ws_情報.Range("B11").Value = lng_Excel列数上限
    
    ' 使用可能なAPIの判定
    ws_情報.Range("A13").Value = "使用可能なAPI:"
    
    str_使用可能API = ""
    
    ' FileSystemObjectの確認
    On Error Resume Next
    Dim obj_FSO As Object
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        str_使用可能API = str_使用可能API & "FileSystemObject, "
    End If
    
    ' WinHTTPの確認
    Dim obj_HTTP As Object
    Set obj_HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number = 0 Then
        str_使用可能API = str_使用可能API & "WinHttp, "
    End If
    
    ' ADOの確認
    Dim obj_接続 As Object
    Set obj_接続 = CreateObject("ADODB.Connection")
    If Err.Number = 0 Then
        str_使用可能API = str_使用可能API & "ADO, "
    End If
    
    ' Outlookの確認
    Dim obj_Outlook As Object
    Set obj_Outlook = CreateObject("Outlook.Application")
    If Err.Number = 0 Then
        str_使用可能API = str_使用可能API & "Outlook, "
    End If
    
    ' XMLの確認
    Dim obj_XML As Object
    Set obj_XML = CreateObject("MSXML2.DOMDocument.6.0")
    If Err.Number = 0 Then
        str_使用可能API = str_使用可能API & "XML DOM, "
    End If
    
    On Error GoTo ErrorHandler
    
    ' 最後のカンマとスペースを削除
    If Len(str_使用可能API) > 2 Then
        str_使用可能API = Left(str_使用可能API, Len(str_使用可能API) - 2)
    End If
    
    ws_情報.Range("B13").Value = str_使用可能API
    
    ' セキュリティ設定
    ws_情報.Range("A15").Value = "マクロセキュリティ:"
    
    ' Accessレベルを取得（間接的な方法）
    Dim str_セキュリティレベル As String
    
    On Error Resume Next
    ' VBAプロジェクトへのアクセスを試みる
    Dim VBProj As Object
    Set VBProj = ThisWorkbook.VBProject
    
    If Err.Number = 0 Then
        str_セキュリティレベル = "VBAプロジェクトにアクセス可能"
    Else
        str_セキュリティレベル = "VBAプロジェクトアクセス制限あり"
    End If
    On Error GoTo ErrorHandler
    
    ws_情報.Range("B15").Value = str_セキュリティレベル
    
    ' ===== ここから環境適応コード例 =====
    ws_情報.Range("A18").Value = "環境適応コード例:"
    
    ' 環境判定後の処理例を追加
    If bln_MAC環境 Then
        ' macOS用のコード
        ws_情報.Range("B18").Value = "macOS環境では、一部のWindowsネイティブAPIが使用できません。"
        ws_情報.Range("B19").Value = "代替手法: AppleScriptやMac固有のオートメーションを検討してください。"
    ElseIf Not bln_64ビット環境 Then
        ' 32bit Windows用のコード
        ws_情報.Range("B18").Value = "32bit環境では、メモリ制限にご注意ください(最大約2GB)。"
        ws_情報.Range("B19").Value = "大量データ処理時はバッチ処理や分割処理を検討してください。"
    Else
        ' 64bit Windows用のコード
        ws_情報.Range("B18").Value = "64bit環境では、Windows APIの宣言に PtrSafe を使用してください。"
        ws_情報.Range("B19").Value = "例: Private Declare PtrSafe Function GetUserName Lib ""advapi32.dll""..."
    End If
    
    ' バージョン依存のコード例を追加
    Dim dbl_バージョン As Double
    dbl_バージョン = CDbl(Application.Version)
    
    ws_情報.Range("A21").Value = "バージョン依存コード例:"
    
    If dbl_バージョン >= 16 Then
        ' Excel 2016以降用のコード
        ws_情報.Range("B21").Value = "Excel 2016以降では、IFS関数やTEXTJOIN関数などの新関数が使用可能です。"
    ElseIf dbl_バージョン >= 14 Then
        ' Excel 2010/2013用のコード
        ws_情報.Range("B21").Value = "Excel 2010/2013では、スライサーやPower Queryを使用することができます。"
    Else
        ' 古いバージョン用のコード
        ws_情報.Range("B21").Value = "このバージョンでは、新しい関数やRibbonUIのカスタマイズに制限があります。"
    End If
    
    ' 書式設定
    ws_情報.Columns("A:B").AutoFit
    ws_情報.Range("A1:B21").Borders.LineStyle = xlContinuous
    
    ' 環境情報をまとめたメッセージを表示
    str_実行環境情報 = "実行環境: " & IIf(bln_64ビット環境, "64ビット", "32ビット") & " Excel " & Application.Version & vbCrLf & _
                     "OS: " & str_OSバージョン & vbCrLf & _
                     "言語: " & str_Office言語
    
    MsgBox str_実行環境情報 & vbCrLf & vbCrLf & _
           "環境に適応したコード例を「環境情報」シートに表示しました。", vbInformation
    
    ws_情報.Activate
    ws_情報.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

'''---------------------------------------------------------
' 2. ユーザー入力エラーの予測と防止策
'''---------------------------------------------------------
Sub ユーザー入力エラー防止()
    ' 変数宣言
    Dim ws_フォーム As Worksheet
    Dim ws_管理 As Worksheet
    Dim rng_入力範囲 As Range
    Dim rng_セル As Range
    Dim arr_電力会社 As Variant
    Dim arr_契約種別 As Variant
    Dim str_メッセージ As String
    
    ' ワークシート設定
    On Error Resume Next
    Set ws_フォーム = ThisWorkbook.Worksheets("入力フォーム")
    Set ws_管理 = ThisWorkbook.Worksheets("管理")
    
    If ws_フォーム Is Nothing Then
        ' 入力フォームシートが存在しない場合、新しく作成する
        Set ws_フォーム = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_フォーム.Name = "入力フォーム"
    End If
    
    If ws_管理 Is Nothing Then
        ' 管理シートが存在しない場合、新しく作成する
        Set ws_管理 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_管理.Name = "管理"
        
        ' 管理シートに初期データを設定
        ws_管理.Range("A1").Value = "電力会社リスト"
        ws_管理.Range("A2").Value = "東京電力"
        ws_管理.Range("A3").Value = "関西電力"
        ws_管理.Range("A4").Value = "中部電力"
        ws_管理.Range("A5").Value = "九州電力"
        ws_管理.Range("A6").Value = "東北電力"
        
        ws_管理.Range("B1").Value = "契約種別リスト"
        ws_管理.Range("B2").Value = "従量電灯A"
        ws_管理.Range("B3").Value = "従量電灯B"
        ws_管理.Range("B4").Value = "従量電灯C"
        ws_管理.Range("B5").Value = "低圧電力"
    End If
    On Error GoTo ErrorHandler
    
    ' 入力フォームの作成
    ws_フォーム.Cells.Clear
    
    ' タイトルと説明
    ws_フォーム.Range("A1").Value = "電力使用量入力フォーム"
    ws_フォーム.Range("A2").Value = "以下の項目を入力してください。入力規則に従った値のみ受け付けます。"
    
    ' 入力フィールドのラベル
    ws_フォーム.Range("A4").Value = "お客様名:"
    ws_フォーム.Range("A5").Value = "電力会社:"
    ws_フォーム.Range("A6").Value = "契約種別:"
    ws_フォーム.Range("A7").Value = "契約アンペア数:"
    ws_フォーム.Range("A8").Value = "使用量(kWh):"
    
    ' 入力セルの書式設定
    ws_フォーム.Range("B4:B8").Interior.Color = RGB(255, 255, 200)
    ws_フォーム.Range("B4:B8").Borders.LineStyle = xlContinuous
    
    ' 電力会社のドロップダウンリスト作成
    arr_電力会社 = ws_管理.Range("A2:A6").Value
    With ws_フォーム.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Join(Application.Transpose(arr_電力会社), ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "電力会社の選択"
        .ErrorTitle = "入力エラー"
        .InputMessage = "リストから電力会社を選択してください"
        .ErrorMessage = "リストにある電力会社を選択してください"
    End With
    
    ' 契約種別のドロップダウンリスト作成
    arr_契約種別 = ws_管理.Range("B2:B5").Value
    With ws_フォーム.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Join(Application.Transpose(arr_契約種別), ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "契約種別の選択"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "リストにある契約種別を選択してください"
    End With
    
    ' アンペア数の入力規則（数値のみ、10〜60の範囲で10刻み）
    With ws_フォーム.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="10,20,30,40,50,60"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "契約アンペア数"
        .ErrorTitle = "入力エラー"
        .InputMessage = "10〜60Aの範囲で10A刻みの値を選択してください"
        .ErrorMessage = "10, 20, 30, 40, 50, 60のいずれかを選択してください"
    End With
    
    ' 使用量の入力規則（正の数値のみ）
    With ws_フォーム.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreater, Formula1:="0"
        .IgnoreBlank = True
        .InputTitle = "使用量入力"
        .ErrorTitle = "入力エラー"
        .InputMessage = "kWh単位で正の数値を入力してください"
        .ErrorMessage = "0より大きい数値を入力してください"
    End With
    
    ' 書式設定
    ws_フォーム.Range("A1").Font.Size = 14
    ws_フォーム.Range("A1").Font.Bold = True
    ws_フォーム.Range("A4:A8").Font.Bold = True
    ws_フォーム.Columns("A:B").AutoFit
    
    ' 送信ボタンの追加
    ws_フォーム.Buttons.Add(200, 200, 100, 30).Select
    Selection.OnAction = "検証して送信"
    Selection.Characters.Text = "検証して送信"
    
    ' フォームを表示
    ws_フォーム.Activate
    ws_フォーム.Range("B4").Select
    
    MsgBox "入力フォームを作成しました。" & vbCrLf & _
           "入力規則を設定し、ユーザー入力エラーを防止しています。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub
