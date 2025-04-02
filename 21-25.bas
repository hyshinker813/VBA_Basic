'''---------------------------------------------------------
' 1. 定型メール送信の自動化（テンプレート活用）
'''---------------------------------------------------------
Sub 定型メール送信自動化()
    ' 変数宣言
    Dim obj_Outlook As Object
    Dim obj_Mail As Object
    Dim ws_顧客 As Worksheet
    Dim lng_最終行 As Long
    Dim lng_行 As Long
    Dim str_テンプレート As String
    Dim str_本文 As String
    Dim str_宛先 As String
    Dim str_宛先名 As String
    Dim str_件名 As String
    Dim str_添付ファイル As String
    Dim str_CC As String
    Dim lng_送信数 As Long
    Dim str_用件 As String ' 未宣言変数を追加
    
    ' Outlookが起動しているか確認
    On Error Resume Next
    Set obj_Outlook = GetObject(, "Outlook.Application")
    If obj_Outlook Is Nothing Then
        Set obj_Outlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' 顧客シートの設定
    Set ws_顧客 = ThisWorkbook.Worksheets("顧客データ")
    
    ' データの最終行を取得
    lng_最終行 = ws_顧客.Cells(ws_顧客.Rows.Count, "A").End(xlUp).Row
    
    ' 送信データがない場合
    If lng_最終行 <= 1 Then
        MsgBox "送信先データがありません。", vbExclamation
        Exit Sub
    End If
    
    ' 確認メッセージ
    If MsgBox("合計 " & lng_最終行 - 1 & " 件のメールを送信します。よろしいですか？", _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If
    
    ' テンプレートの読み込み（例として組み込みテンプレート）
    str_テンプレート = "拝啓 [顧客名] 様" & vbCrLf & vbCrLf & _
                      "いつもお世話になっております。" & vbCrLf & _
                      "[会社名]の[担当者]です。" & vbCrLf & vbCrLf & _
                      "本日は下記の件につきましてご連絡いたしました。" & vbCrLf & _
                      "・[用件]" & vbCrLf & vbCrLf & _
                      "詳細は添付ファイルをご確認ください。" & vbCrLf & _
                      "ご不明な点がございましたら、お気軽にお問い合わせください。" & vbCrLf & vbCrLf & _
                      "よろしくお願いいたします。" & vbCrLf & vbCrLf & _
                      "敬具" & vbCrLf & vbCrLf & _
                      "------------------------------" & vbCrLf & _
                      "[会社名]" & vbCrLf & _
                      "[担当者]" & vbCrLf & _
                      "TEL: [電話番号]" & vbCrLf & _
                      "Email: [メールアドレス]" & vbCrLf & _
                      "------------------------------"
    
    ' 自社情報の設定
    str_テンプレート = Replace(str_テンプレート, "[会社名]", "株式会社テプコソリューションアドバンス")
    str_テンプレート = Replace(str_テンプレート, "[担当者]", "電力 太郎")
    str_テンプレート = Replace(str_テンプレート, "[電話番号]", "03-1234-5678")
    str_テンプレート = Replace(str_テンプレート, "[メールアドレス]", "taro.denryoku@example.com")
    
    ' 添付ファイルのパス設定
    str_添付ファイル = ThisWorkbook.Path & "\資料.pdf"
    
    ' 添付ファイルの存在確認
    If Dir(str_添付ファイル) = "" Then
        If MsgBox("添付ファイルが見つかりません: " & str_添付ファイル & vbCrLf & _
                 "添付なしで続行しますか？", vbQuestion + vbYesNo, "確認") = vbNo Then
            Exit Sub
        End If
        str_添付ファイル = ""
    End If
    
    ' 進捗表示の初期化
    Application.StatusBar = "メール送信準備中..."
    
    ' 送信数の初期化
    lng_送信数 = 0
    
    ' 各顧客にメールを送信
    For lng_行 = 2 To lng_最終行 ' ヘッダー行をスキップ
        ' データの取得
        str_宛先 = ws_顧客.Cells(lng_行, 3).Value ' メールアドレス列
        str_宛先名 = ws_顧客.Cells(lng_行, 2).Value ' 顧客名列
        str_用件 = ws_顧客.Cells(lng_行, 4).Value ' 用件列
        
        ' 必須データのチェック
        If Trim(str_宛先) = "" Then
            ' メールアドレスがない場合はスキップ
            ws_顧客.Cells(lng_行, 5).Value = "エラー: メールアドレスなし"
            GoTo NextCustomer
        End If
        
        ' 進捗表示を更新
        Application.StatusBar = "メール送信中... " & lng_行 - 1 & "/" & lng_最終行 - 1 & _
                              " (" & Format((lng_行 - 1) / (lng_最終行 - 1), "0%") & ")"
        
        ' 本文のカスタマイズ
        str_本文 = str_テンプレート
        str_本文 = Replace(str_本文, "[顧客名]", str_宛先名)
        str_本文 = Replace(str_本文, "[用件]", str_用件)
        
        ' 件名の設定
        str_件名 = "【ご連絡】" & str_用件 & "について"
        
        ' メール作成
        Set obj_Mail = obj_Outlook.CreateItem(0) ' olMailItem
        
        With obj_Mail
            .To = str_宛先
            ' CC設定（オプション）
            If ws_顧客.Cells(lng_行, 6).Value <> "" Then
                .CC = ws_顧客.Cells(lng_行, 6).Value
            End If
            .Subject = str_件名
            .Body = str_本文
            
            ' 添付ファイルの追加（ファイルが存在する場合）
            If str_添付ファイル <> "" Then
                .Attachments.Add str_添付ファイル
            End If
            
            ' メールの送信（または下書き保存、表示など）
            ' コメントアウトを切り替えて目的の動作を選択
            
            ' 方法1: 直接送信（自動送信）
            '.Send
            
            ' 方法2: 下書きとして保存
            '.Save
            
            ' 方法3: メールを表示（ユーザーが手動で確認・送信）
            .Display
        End With
        
        ' 送信ログの記録
        ws_顧客.Cells(lng_行, 5).Value = "送信済: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
        
        ' 送信数をカウント
        lng_送信数 = lng_送信数 + 1
        
        ' 少し待機（サーバー負荷軽減のため）
        Application.Wait Now + TimeSerial(0, 0, 1)
        
NextCustomer:
    Next lng_行
    
    ' 正常終了処理
    Application.StatusBar = False
    MsgBox "メール送信処理が完了しました。" & vbCrLf & _
           "送信数: " & lng_送信数 & "/" & lng_最終行 - 1, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
'''---------------------------------------------------------
' 4. 受信メールのフィルタリングと処理
'''---------------------------------------------------------
Sub 受信メールフィルタリング処理()
    ' 変数宣言
    Dim obj_Outlook As Object
    Dim obj_Namespace As Object
    Dim obj_受信トレイ As Object
    Dim obj_メール As Object
    Dim obj_フォルダ As Object
    Dim col_メール As Object
    Dim lng_合計 As Long
    Dim lng_処理数 As Long
    Dim i As Long
    
    ' Outlookが起動しているか確認
    On Error Resume Next
    Set obj_Outlook = GetObject(, "Outlook.Application")
    If obj_Outlook Is Nothing Then
        Set obj_Outlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' Outlookオブジェクトの取得
    Set obj_Namespace = obj_Outlook.GetNamespace("MAPI")
    
    ' 受信トレイの取得
    Set obj_受信トレイ = obj_Namespace.GetDefaultFolder(6) ' olFolderInbox
    
    ' フィルタリング条件の設定
    ' 例：未読で、特定の送信者からのメールを検索
    Dim str_フィルタ As String
    str_フィルタ = "[Unread] = True AND " & _
                  "[SenderEmailAddress] = 'important@example.com'"
    
    ' フィルタに一致するメールを取得
    Set col_メール = obj_受信トレイ.Items.Restrict(str_フィルタ)
    
    ' メールの日付でソート（新しい順）
    col_メール.Sort "[ReceivedTime]", True
    
    ' メールの合計数
    lng_合計 = col_メール.Count
    
    ' 処理するメールがない場合
    If lng_合計 = 0 Then
        MsgBox "条件に一致するメールがありません。", vbInformation
        Exit Sub
    End If
    
    ' 確認メッセージ
    If MsgBox("条件に一致するメール " & lng_合計 & " 件を処理します。よろしいですか？", _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If
    
    ' 処理結果を記録するシートを準備
    Dim ws_結果 As Worksheet
    
    On Error Resume Next
    Set ws_結果 = ThisWorkbook.Worksheets("メール処理結果")
    If ws_結果 Is Nothing Then
        Set ws_結果 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws_結果.Name = "メール処理結果"
    End If
    ws_結果.Cells.Clear
    On Error GoTo ErrorHandler
    
    ' ヘッダー行の設定
    ws_結果.Range("A1").Value = "処理日時"
    ws_結果.Range("B1").Value = "件名"
    ws_結果.Range("C1").Value = "送信者"
    ws_結果.Range("D1").Value = "受信日時"
    ws_結果.Range("E1").Value = "重要度"
    ws_結果.Range("F1").Value = "処理結果"
    ws_結果.Range("A1:F1").Font.Bold = True
    
    ' 処理数の初期化
    lng_処理数 = 0
    
    ' 各メールを処理
    For i = 1 To lng_合計
        ' メールオブジェクトの取得
        Set obj_メール = col_メール.Item(i)
        
        ' 進捗表示を更新
        Application.StatusBar = "メール処理中... " & i & "/" & lng_合計 & _
                              " (" & Format(i / lng_合計, "0%") & ")"
        
        ' メール情報の取得
        Dim str_件名 As String
        Dim str_送信者 As String
        Dim dt_受信日時 As Date
        Dim int_重要度 As Integer
        
        str_件名 = obj_メール.Subject
        str_送信者 = obj_メール.SenderName & " <" & obj_メール.SenderEmailAddress & ">"
        dt_受信日時 = obj_メール.ReceivedTime
        int_重要度 = obj_メール.Importance ' 1=低, 2=標準, 3=高
        
        ' メール処理のロジック（例：特定のフォルダに移動）
        Dim str_処理結果 As String
        str_処理結果 = ""
        
        ' 例1: 特定の件名を含むメールを処理
        If InStr(1, str_件名, "報告", vbTextCompare) > 0 Then
            ' 「報告」フォルダに移動
            On Error Resume Next
            Set obj_フォルダ = obj_受信トレイ.Folders("報告")
            If obj_フォルダ Is Nothing Then
                ' フォルダが存在しない場合は作成
                Set obj_フォルダ = obj_受信トレイ.Folders.Add("報告")
            End If
            
            ' メールを移動
            obj_メール.Move obj_フォルダ
            str_処理結果 = "「報告」フォルダに移動"
            On Error GoTo ErrorHandler
        
        ' 例2: 重要度の高いメールを処理
        ElseIf int_重要度 = 3 Then
            ' 高重要度フラグを設定
            obj_メール.FlagStatus = 2 ' olFlagMarked
            str_処理結果 = "高重要度フラグを設定"
        
        ' 例3: 過去のメールをアーカイブ
        ElseIf dt_受信日時 < Date - 30 Then
            ' 「アーカイブ」フォルダに移動
            On Error Resume Next
            Set obj_フォルダ = obj_受信トレイ.Folders("アーカイブ")
            If obj_フォルダ Is Nothing Then
                Set obj_フォルダ = obj_受信トレイ.Folders.Add("アーカイブ")
            End If
            
            ' メールを移動
            obj_メール.Move obj_フォルダ
            str_処理結果 = "「アーカイブ」フォルダに移動"
            On Error GoTo ErrorHandler
        
        ' 例4: その他のメールを既読にするだけ
        Else
            obj_メール.UnRead = False
            str_処理結果 = "既読に設定"
        End If
        
        ' 結果をシートに記録
        ws_結果.Cells(lng_処理数 + 2, 1).Value = Now
        ws_結果.Cells(lng_処理数 + 2, 2).Value = str_件名
        ws_結果.Cells(lng_処理数 + 2, 3).Value = str_送信者
        ws_結果.Cells(lng_処理数 + 2, 4).Value = dt_受信日時
        ws_結果.Cells(lng_処理数 + 2, 5).Value = Choose(int_重要度, "低", "標準", "高")
        ws_結果.Cells(lng_処理数 + 2, 6).Value = str_処理結果
        
        ' 処理数をカウント
        lng_処理数 = lng_処理数 + 1
    Next i
    
    ' 書式設定
    ws_結果.Columns("A:F").AutoFit
    ws_結果.Range("A2:F" & lng_処理数 + 1).Borders.LineStyle = xlContinuous
    ws_結果.Range("A1:F" & lng_処理数 + 1).Sort Key1:=ws_結果.Range("D1"), Order1:=xlDescending, Header:=xlYes
    
    ' 正常終了処理
    Application.StatusBar = False
    ws_結果.Activate
    
    MsgBox "メールフィルタリング処理が完了しました。" & vbCrLf & _
           "処理メール数: " & lng_処理数 & " / " & lng_合計, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
'''---------------------------------------------------------
' 5. メール本文のHTML整形と装飾
'''---------------------------------------------------------
Sub メール本文HTML整形装飾()
    ' 変数宣言
    Dim obj_Outlook As Object
    Dim obj_Mail As Object
    Dim ws_データ As Worksheet
    Dim str_HTML本文 As String
    Dim str_CSS As String
    Dim str_ヘッダー As String
    Dim str_フッター As String
    Dim str_テーブル As String
    Dim lng_最終行 As Long
    Dim i As Long
    
    ' Outlookが起動しているか確認
    On Error Resume Next
    Set obj_Outlook = GetObject(, "Outlook.Application")
    If obj_Outlook Is Nothing Then
        Set obj_Outlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' ワークシートの設定
    Set ws_データ = ThisWorkbook.Worksheets("レポートデータ")
    
    ' データの最終行を取得
    lng_最終行 = ws_データ.Cells(ws_データ.Rows.Count, "A").End(xlUp).Row
    
    ' データがない場合
    If lng_最終行 <= 1 Then
        MsgBox "レポートデータがありません。", vbExclamation
        Exit Sub
    End If
    
    ' ===== CSS スタイルの定義 =====
    str_CSS = "<style>" & _
             "body { font-family: Arial, sans-serif; margin: 0; padding: 20px; color: #333; }" & _
             "h1 { color: #003366; border-bottom: 2px solid #003366; padding-bottom: 10px; }" & _
             "h2 { color: #003366; margin-top: 20px; }" & _
             "table { border-collapse: collapse; width: 100%; margin-top: 20px; }" & _
             "th { background-color: #003366; color: white; text-align: left; padding: 8px; }" & _
             "td { border: 1px solid #ddd; padding: 8px; }" & _
             "tr:nth-child(even) { background-color: #f2f2f2; }" & _
             ".highlight { background-color: #ffffcc; }" & _
             ".footer { margin-top: 30px; border-top: 1px solid #ddd; padding-top: 10px; font-size: 85%; color: #777; }" & _
             ".chart { margin: 20px 0; text-align: center; }" & _
             "</style>"
    
    ' ===== HTML ヘッダー部分 =====
    str_ヘッダー = "<html>" & _
                 "<head>" & _
                 "<meta charset='utf-8'>" & _
                 str_CSS & _
                 "</head>" & _
                 "<body>" & _
                 "<h1>月次エネルギー使用レポート</h1>" & _
                 "<p>このレポートは " & Format(Date, "yyyy年m月d日") & " に作成されました。</p>" & _
                 "<p>以下に、各拠点のエネルギー使用状況をまとめています。</p>"
    
    ' ===== テーブルの作成 =====
    str_テーブル = "<h2>拠点別エネルギー使用量</h2>" & _
                 "<table>" & _
                 "<tr>" & _
                 "<th>拠点名</th>" & _
                 "<th>電力使用量 (kWh)</th>" & _
                 "<th>ガス使用量 (m³)</th>" & _
                 "<th>前月比</th>" & _
                 "<th>状態</th>" & _
                 "</tr>"
    
    ' テーブルの行を追加
    For i = 2 To lng_最終行 ' ヘッダー行をスキップ
        Dim str_拠点名 As String
        Dim dbl_電力使用量 As Double
        Dim dbl_ガス使用量 As Double
        Dim dbl_前月比 As Double
        Dim str_状態 As String
        Dim str_行スタイル As String
        
        ' データ取得
        str_拠点名 = ws_データ.Cells(i, 1).Value
        dbl_電力使用量 = ws_データ.Cells(i, 2).Value
        dbl_ガス使用量 = ws_データ.Cells(i, 3).Value
        dbl_前月比 = ws_データ.Cells(i, 4).Value
        
        ' 状態の判定（例：前月比がプラスなら「増加」、マイナスなら「削減」）
        If dbl_前月比 > 0 Then
            str_状態 = "増加"
            str_行スタイル = " class='highlight'"  ' 増加した行を強調表示
        ElseIf dbl_前月比 < 0 Then
            str_状態 = "削減"
            str_行スタイル = ""
        Else
            str_状態 = "変化なし"
            str_行スタイル = ""
        End If
        
        ' テーブル行の追加
        str_テーブル = str_テーブル & _
                       "<tr" & str_行スタイル & ">" & _
                       "<td>" & str_拠点名 & "</td>" & _
                       "<td>" & Format(dbl_電力使用量, "#,##0") & "</td>" & _
                       "<td>" & Format(dbl_ガス使用量, "#,##0") & "</td>" & _
                       "<td>" & Format(dbl_前月比, "+0.0%;-0.0%;0.0%") & "</td>" & _
                       "<td>" & str_状態 & "</td>" & _
                       "</tr>"
    Next i
    
    ' テーブルを閉じる
    str_テーブル = str_テーブル & "</table>"
    
    ' ===== グラフのプレースホルダー（実際のグラフはOutlookでは表示できないが、イメージとして表現） =====
    Dim str_グラフ As String
    str_グラフ = "<div class='chart'>" & _
                "<h2>月次推移グラフ</h2>" & _
                "<p>[このメールではグラフを表示できません。添付ファイルまたはWebダッシュボードをご確認ください。]</p>" & _
                "</div>"
    
    ' ===== フッター部分 =====
    str_フッター = "<div class='footer'>" & _
                 "<p>このレポートは自動生成されています。詳細な分析は添付ファイルをご確認ください。</p>" & _
                 "<p>ご質問がございましたら、以下の連絡先までお問い合わせください。</p>" & _
                 "<p>株式会社テプコソリューションアドバンス<br>" & _
                 "エネルギー管理部<br>" & _
                 "TEL: 03-1234-5678<br>" & _
                 "Email: energy@example.com</p>" & _
                 "</div>" & _
                 "</body>" & _
                 "</html>"
    
    ' ===== 完全なHTML本文を構築 =====
    str_HTML本文 = str_ヘッダー & str_テーブル & str_グラフ & str_フッター
    
    ' Outlookメールの作成
    Set obj_Mail = obj_Outlook.CreateItem(0) ' olMailItem
    
    With obj_Mail
        ' 宛先の設定
        .To = "recipient@example.com"
        
        ' CC/BCC設定（オプション）
        .CC = "manager@example.com"
        
        ' 件名設定
        .Subject = "月次エネルギー使用レポート - " & Format(Date, "yyyy年m月")
        
        ' HTML形式の本文を設定
        .HTMLBody = str_HTML本文
        
        ' 添付ファイルの追加（オプション）
        Dim str_添付ファイル As String
        str_添付ファイル = ThisWorkbook.Path & "\エネルギーレポート_" & Format(Date, "yyyymm") & ".xlsx"
        
        ' 添付ファイルが存在する場合のみ添付
        If Dir(str_添付ファイル) <> "" Then
            .Attachments.Add str_添付ファイル
        End If
        
        ' メールの表示（確認用）
        .Display
        
        ' 直接送信する場合（コメントアウトを解除して使用）
        '.Send
    End With
    
    MsgBox "HTML形式のレポートメールを作成しました。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
End Sub

'''---------------------------------------------------------
' 2. メール添付ファイルの自動保存と処理
'''---------------------------------------------------------
Sub メール添付ファイル自動保存()
    ' 変数宣言
    Dim obj_Outlook As Object
    Dim obj_Namespace As Object
    Dim obj_受信トレイ As Object
    Dim obj_フォルダ As Object
    Dim obj_メール As Object
    Dim obj_添付ファイル As Object
    Dim str_保存フォルダ As String
    Dim str_保存パス As String
    Dim str_ファイル名 As String
    Dim str_フィルタ As String
    Dim lng_処理数 As Long
    Dim lng_合計 As Long
    Dim i As Long
    Dim obj_FSO As Object
    
    ' FSO（FileSystemObject）の作成
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Outlookが起動しているか確認
    On Error Resume Next
    Set obj_Outlook = GetObject(, "Outlook.Application")
    If obj_Outlook Is Nothing Then
        Set obj_Outlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' 保存先フォルダの設定
    str_保存フォルダ = ThisWorkbook.Path & "\添付ファイル\"
    
    ' フォルダが存在しない場合は作成
    If Not obj_FSO.FolderExists(str_保存フォルダ) Then
        obj_FSO.CreateFolder str_保存フォルダ
    End If
    
    ' Outlookオブジェクトの取得
    Set obj_Namespace = obj_Outlook.GetNamespace("MAPI")
    
    ' 受信トレイの取得（または他のフォルダ）
    Set obj_受信トレイ = obj_Namespace.GetDefaultFolder(6) ' olFolderInbox
    
    ' 特定のサブフォルダを使用する場合（オプション）
    'Set obj_フォルダ = obj_受信トレイ.Folders("プロジェクトA")
    Set obj_フォルダ = obj_受信トレイ ' またはサブフォルダ
    
    ' メールのフィルタ条件（オプション）
    str_フィルタ = "[Unread] = True" ' 未読メールのみ
    'str_フィルタ = "[Subject] LIKE '%レポート%'" ' 件名に特定の文字列を含むメール
    'str_フィルタ = "[ReceivedTime] > '" & Format(Date - 7, "mm/dd/yyyy") & "'" ' 過去7日以内のメール
    
    ' フィルタに一致するメールを取得
    Dim col_メール As Object
    Set col_メール = obj_フォルダ.Items
    
    ' フィルタを適用（空の場合はすべてのメール）
    If str_フィルタ <> "" Then
        col_メール = col_メール.Restrict(str_フィルタ)
    End If
    
    ' メールの合計数
    lng_合計 = col_メール.Count
    
    ' 処理するメールがない場合
    If lng_合計 = 0 Then
        MsgBox "処理対象のメールがありません。", vbInformation
        Exit Sub
    End If
    
    ' 確認メッセージ
    If MsgBox("合計 " & lng_合計 & " 件のメールを処理します。よろしいですか？", _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If
    
    ' 処理数の初期化
    lng_処理数 = 0
    
    ' 日付フォルダの作成（オプション - 日付ごとに整理する場合）
    Dim str_日付フォルダ As String
    str_日付フォルダ = str_保存フォルダ & Format(Date, "yyyymmdd") & "\"
    
    If Not obj_FSO.FolderExists(str_日付フォルダ) Then
        obj_FSO.CreateFolder str_日付フォルダ
    End If
    
    ' アクセスログの記録開始
    Dim str_ログファイル As String
    Dim int_ファイル番号 As Integer
    
    str_ログファイル = str_日付フォルダ & "添付ファイル_ログ_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    int_ファイル番号 = FreeFile
    
    Open str_ログファイル For Output As #int_ファイル番号
    Print #int_ファイル番号, "添付ファイル処理ログ - " & Format(Now, "yyyy/mm/dd hh:nn:ss")
    Print #int_ファイル番号, "------------------------------------------------------"
    
    ' 各メールを処理
    For i = 1 To lng_合計
        ' 通常は新しいものから処理（最新が一番上）
        Set obj_メール = col_メール.Item(i)
        
        ' 進捗表示を更新
        Application.StatusBar = "メール処理中... " & i & "/" & lng_合計 & _
                              " (" & Format(i / lng_合計, "0%") & ")"
        
        ' 添付ファイルがあるか確認
        If obj_メール.Attachments.Count > 0 Then
            ' ログにメール情報を記録
            Print #int_ファイル番号, "メール件名: " & obj_メール.Subject
            Print #int_ファイル番号, "送信者: " & obj_メール.SenderName & " <" & obj_メール.SenderEmailAddress & ">"
            Print #int_ファイル番号, "受信日時: " & obj_メール.ReceivedTime
            Print #int_ファイル番号, "添付ファイル数: " & obj_メール.Attachments.Count
            
            ' 各添付ファイルを処理
            For Each obj_添付ファイル In obj_メール.Attachments
                ' ファイル名を取得
                str_ファイル名 = obj_添付ファイル.FileName
                
                ' 特定の拡張子のみ処理（オプション）
                Dim str_拡張子 As String
                str_拡張子 = LCase(Right(str_ファイル名, 4))
                
                If str_拡張子 = ".xls" Or str_拡張子 = "xlsx" Or str_拡張子 = ".pdf" Or str_拡張子 = ".csv" Then
                    ' 保存パスの設定（重複を避けるため送信者名とタイムスタンプを追加）
                    Dim str_タイムスタンプ As String
                    str_タイムスタンプ = Format(Now, "yyyymmdd_hhnnss") & "_" & i
                    
                    ' ファイル名の安全対策（無効な文字を除去）
                    str_ファイル名 = Replace(Replace(Replace(Replace(str_ファイル名, "\", "_"), "/", "_"), ":", "_"), "*", "_")
                    
                    ' 最終的な保存パス
                    str_保存パス = str_日付フォルダ & str_タイムスタンプ & "_" & str_ファイル名
                    
                    ' ファイルを保存
                    obj_添付ファイル.SaveAsFile str_保存パス
                    
                    ' ログに記録
                    Print #int_ファイル番号, "  - 保存: " & str_ファイル名 & " -> " & str_保存パス
                    
                    ' Excel添付ファイルを自動的に分析する（オプション）
                    If str_拡張子 = ".xls" Or str_拡張子 = "xlsx" Then
                        ' ここに分析ロジックを入れることができます
                        ' 例：
                        ' Call Excel添付ファイル分析(str_保存パス)
                    End If
                    
                    ' 処理数をカウント
                    lng_処理数 = lng_処理数 + 1
                Else
                    ' 対象外の拡張子
                    Print #int_ファイル番号, "  - スキップ: " & str_ファイル名 & " (対象外の形式)"
                End If
            Next obj_添付ファイル
            
            ' メールにフラグを設定（オプション）
            obj_メール.FlagStatus = 1 ' olFlagComplete
            
            ' メールを既読にする（オプション）
            obj_メール.UnRead = False
            
            ' メールをフォルダ移動（オプション）
            'Dim obj_処理済フォルダ As Object
            'Set obj_処理済フォルダ = obj_受信トレイ.Folders("処理済")
            'If Not obj_処理済フォルダ Is Nothing Then
            '    obj_メール.Move obj_処理済フォルダ
            'End If
            
            Print #int_ファイル番号, "------------------------------------------------------"
        End If
    Next i
    
    ' ログファイルを閉じる
    Print #int_ファイル番号, ""
    Print #int_ファイル番号, "処理完了"
    Print #int_ファイル番号, "合計処理ファイル数: " & lng_処理数
    Print #int_ファイル番号, "処理終了時刻: " & Format(Now, "yyyy/mm/dd hh:nn:ss")
    Close #int_ファイル番号
    
    ' 正常終了処理
    Application.StatusBar = False
    MsgBox "添付ファイルの処理が完了しました。" & vbCrLf & _
           "処理ファイル数: " & lng_処理数 & vbCrLf & _
           "保存先: " & str_日付フォルダ, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    
    ' ファイルハンドルが開いている場合は閉じる
    On Error Resume Next
    Close #int_ファイル番号
    
    Application.StatusBar = False
End Sub

'''---------------------------------------------------------
' 3. エクセルデータを元にしたメール配信
'''---------------------------------------------------------
Sub エクセルデータメール配信()
    ' 変数宣言
    Dim obj_Outlook As Object
    Dim obj_Mail As Object
    Dim ws_データ As Worksheet
    Dim ws_テンプレート As Worksheet
    Dim rng_表 As ListObject
    Dim rng_行 As ListRow
    Dim str_HTML本文 As String
    Dim str_テンプレート As String
    Dim str_置換前 As String
    Dim str_置換後 As String
    Dim lng_送信数 As Long
    Dim bln_HTML形式 As Boolean
    Dim rng_新列 As ListColumn ' 未宣言変数を追加
    Dim i As Long ' 未宣言変数を追加
    
    ' HTML形式かプレーンテキスト形式かの設定
    bln_HTML形式 = True
    
    ' Outlookが起動しているか確認
    On Error Resume Next
    Set obj_Outlook = GetObject(, "Outlook.Application")
    If obj_Outlook Is Nothing Then
        Set obj_Outlook = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' ワークシートの設定
    Set ws_データ = ThisWorkbook.Worksheets("送信データ")
    Set ws_テンプレート = ThisWorkbook.Worksheets("メールテンプレート")
    
    ' データが表（ListObject）形式かチェック
    If ws_データ.ListObjects.Count = 0 Then
        MsgBox "送信データが表形式になっていません。" & vbCrLf & _
               "データを表に変換してから再実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' 送信データ表の取得
    Set rng_表 = ws_データ.ListObjects(1)
    
    ' 送信データがない場合
    If rng_表.ListRows.Count = 0 Then
        MsgBox "送信データがありません。", vbExclamation
        Exit Sub
    End If
    
    ' 必須列の確認（例：メールアドレス列）
    Dim lng_メールアドレス列 As Long
    Dim lng_送信状態列 As Long
    
    On Error Resume Next
    lng_メールアドレス列 = rng_表.ListColumns("メールアドレス").Index
    lng_送信状態列 = rng_表.ListColumns("送信状態").Index
    On Error GoTo ErrorHandler
    
    If lng_メールアドレス列 = 0 Then
        MsgBox "「メールアドレス」列が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 送信状態列がなければ追加
    If lng_送信状態列 = 0 Then
        Set rng_新列 = rng_表.ListColumns.Add
        rng_新列.Name = "送信状態"
        lng_送信状態列 = rng_表.ListColumns.Count
    End If
    
    ' テンプレートの読み込み
    If bln_HTML形式 Then
        ' HTML形式のテンプレート
        str_テンプレート = ws_テンプレート.Range("B1").Value
        
        ' HTML形式のテンプレートが空の場合、デフォルトのHTMLを使用
        If Trim(str_テンプレート) = "" Then
            str_テンプレート = "<html>" & _
                              "<head>" & _
                              "<style>" & _
                              "body { font-family: Arial, sans-serif; }" & _
                              "h1 { color: #003366; }" & _
                              "table { border-collapse: collapse; }" & _
                              "th, td { border: 1px solid #ddd; padding: 8px; }" & _
                              "th { background-color: #f2f2f2; }" & _
                              "</style>" & _
                              "</head>" & _
                              "<body>" & _
                              "<h1>{{タイトル}}</h1>" & _
                              "<p>{{宛先}}様</p>" & _
                              "<p>いつもお世話になっております。{{会社名}}の{{担当者}}です。</p>" & _
                              "<p>下記の通りご連絡いたします。</p>" & _
                              "<ul>" & _
                              "<li>{{項目1}}</li>" & _
                              "<li>{{項目2}}</li>" & _
                              "</ul>" & _
                              "<p>詳細は以下をご確認ください。</p>" & _
                              "<table>" & _
                              "<tr><th>項目</th><th>内容</th></tr>" & _
                              "<tr><td>{{項目A}}</td><td>{{内容A}}</td></tr>" & _
                              "<tr><td>{{項目B}}</td><td>{{内容B}}</td></tr>" & _
                              "</table>" & _
                              "<p>よろしくお願いいたします。</p>" & _
                              "<hr>" & _
                              "<p>{{会社名}}<br>" & _
                              "{{担当者}}<br>" & _
                              "TEL: {{電話番号}}<br>" & _
                              "Email: {{メールアドレス}}</p>" & _
                              "</body>" & _
                              "</html>"
        End If
    Else
        ' プレーンテキスト形式のテンプレート
        str_テンプレート = ws_テンプレート.Range("A1").Value
        
        ' テキスト形式のテンプレートが空の場合、デフォルトのテキストを使用
        If Trim(str_テンプレート) = "" Then
            str_テンプレート = "{{タイトル}}" & vbCrLf & vbCrLf & _
                              "{{宛先}}様" & vbCrLf & vbCrLf & _
                              "いつもお世話になっております。{{会社名}}の{{担当者}}です。" & vbCrLf & vbCrLf & _
                              "下記の通りご連絡いたします。" & vbCrLf & _
                              "・{{項目1}}" & vbCrLf & _
                              "・{{項目2}}" & vbCrLf & vbCrLf & _
                              "詳細は以下をご確認ください。" & vbCrLf & _
                              "-------------------" & vbCrLf & _
                              "項目: {{項目A}}" & vbCrLf & _
                              "内容: {{内容A}}" & vbCrLf & vbCrLf & _
                              "項目: {{項目B}}" & vbCrLf & _
                              "内容: {{内容B}}" & vbCrLf & _
                              "-------------------" & vbCrLf & vbCrLf & _
                              "よろしくお願いいたします。" & vbCrLf & vbCrLf & _
                              "------------------------------" & vbCrLf & _
                              "{{会社名}}" & vbCrLf & _
                              "{{担当者}}" & vbCrLf & _
                              "TEL: {{電話番号}}" & vbCrLf & _
                              "Email: {{メールアドレス}}" & vbCrLf & _
                              "------------------------------"
        End If
    End If
    
    ' 自社情報の設定
    str_テンプレート = Replace(str_テンプレート, "{{会社名}}", "株式会社テプコソリューションアドバンス")
    str_テンプレート = Replace(str_テンプレート, "{{担当者}}", "電力 太郎")
    str_テンプレート = Replace(str_テンプレート, "{{電話番号}}", "03-1234-5678")
    str_テンプレート = Replace(str_テンプレート, "{{メールアドレス}}", "taro.denryoku@example.com")
    
    ' 確認メッセージ
    If MsgBox("合計 " & rng_表.ListRows.Count & " 件のメールを送信します。よろしいですか？", _
              vbQuestion + vbYesNo, "確認") = vbNo Then
        Exit Sub
    End If
    
    ' 送信数の初期化
    lng_送信数 = 0
    
    ' 進捗表示の初期化
    Application.StatusBar = "メール送信準備中..."
    
    ' 各行のデータでメールを送信
    For Each rng_行 In rng_表.ListRows
        ' 既に送信済みならスキップ（オプション）
        If lng_送信状態列 > 0 Then
            If InStr(1, rng_行.Range(1, lng_送信状態列).Value, "送信済", vbTextCompare) > 0 Then
                GoTo NextRow
            End If
        End If
        
        ' メールアドレスの取得
        Dim str_宛先 As String
        str_宛先 = rng_行.Range(1, lng_メールアドレス列).Value
        
        ' メールアドレスが空ならスキップ
        If Trim(str_宛先) = "" Then
            If lng_送信状態列 > 0 Then
                rng_行.Range(1, lng_送信状態列).Value = "エラー: メールアドレスなし"
            End If
            GoTo NextRow
        End If
        
        ' 進捗表示を更新
        Application.StatusBar = "メール送信中... " & lng_送信数 + 1 & "/" & rng_表.ListRows.Count & _
                              " (" & Format((lng_送信数 + 1) / rng_表.ListRows.Count, "0%") & ")"
        
        ' テンプレートをカスタマイズ
        str_HTML本文 = str_テンプレート
        
        ' 各列のデータでテンプレートのプレースホルダーを置換
        For i = 1 To rng_表.ListColumns.Count
            str_置換前 = "{{" & rng_表.ListColumns(i).Name & "}}"
            str_置換後 = rng_行.Range(1, i).Value
            
            ' 置換
            str_HTML本文 = Replace(str_HTML本文, str_置換前, str_置換後)
        Next i
        
        ' メール作成
        Set obj_Mail = obj_Outlook.CreateItem(0) ' olMailItem
        
        With obj_Mail
            .To = str_宛先
            
            ' CC設定（オプション - 対応する列がある場合）
            On Error Resume Next
            Dim lng_CC列 As Long ' 未宣言変数を追加
            lng_CC列 = rng_表.ListColumns("CC").Index
            If lng_CC列 > 0 Then
                If Trim(rng_行.Range(1, lng_CC列).Value) <> "" Then
                    .CC = rng_行.Range(1, lng_CC列).Value
                End If
            End If
            On Error GoTo ErrorHandler
            
            ' 件名設定
            On Error Resume Next
            Dim lng_件名列 As Long ' 未宣言変数を追加
            lng_件名列 = rng_表.ListColumns("件名").Index
            If lng_件名列 > 0 Then
                .Subject = rng_行.Range(1, lng_件名列).Value
            Else
                ' デフォルト件名
                .Subject = "【お知らせ】重要なお知らせ"
            End If
            On Error GoTo ErrorHandler
            
            ' HTML形式またはテキスト形式で本文を設定
            If bln_HTML形式 Then
                .HTMLBody = str_HTML本文
            Else
                .Body = str_HTML本文
            End If
            
            ' 添付ファイル（オプション - 対応する列がある場合）
            On Error Resume Next
            Dim lng_添付ファイル列 As Long ' 未宣言変数を追加
            lng_添付ファイル列 = rng_表.ListColumns("添付ファイル").Index
            If lng_添付ファイル列 > 0 Then
                If Trim(rng_行.Range(1, lng_添付ファイル列).Value) <> "" Then
                    Dim str_添付ファイルパス As String
                    str_添付ファイルパス = ThisWorkbook.Path & "\" & rng_行.Range(1, lng_添付ファイル列).Value
                    
                    If Dir(str_添付ファイルパス) <> "" Then
                        .Attachments.Add str_添付ファイルパス
                    End If
                End If
            End If
            On Error GoTo ErrorHandler
            
            ' メールの送信または表示
            ' コメントアウトを切り替えて目的の動作を選択
            
            ' 方法1: 直接送信（自動送信）
            '.Send
            
            ' 方法2: 下書きとして保存
            '.Save
            
            ' 方法3: メールを表示（ユーザーが手動で確認・送信）
            .Display
        End With
        
        ' 送信状態を更新
        If lng_送信状態列 > 0 Then
            rng_行.Range(1, lng_送信状態列).Value = "送信済: " & Format(Now, "yyyy/mm/dd hh:mm:ss")
        End If
        
        ' 送信数をカウント
        lng_送信数 = lng_送信数 + 1
        
        ' 少し待機（サーバー負荷軽減のため）
        Application.Wait Now + TimeSerial(0, 0, 1)
        
NextRow:
    Next rng_行
    
    ' 正常終了処理
    Application.StatusBar = False
    MsgBox "メール送信処理が完了しました。" & vbCrLf & _
           "送信数: " & lng_送信数 & "/" & rng_表.ListRows.Count, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
End Sub
