'''---------------------------------------------------------
' 1. 進捗バーの実装と進捗状況表示
'''---------------------------------------------------------
Sub 進捗バー実装()
    ' 変数宣言
    Dim frm_進捗画面 As Object
    Dim ws_処理対象 As Worksheet
    Dim lng_最終行 As Long
    Dim lng_現在行 As Long
    Dim lng_処理総数 As Long
    Dim sng_進捗率 As Single
    Dim str_ステータス As String
    Dim dbl_開始時間 As Double
    Dim dbl_経過時間 As Double
    Dim dbl_推定残り時間 As Double
    
    ' 処理対象のワークシートを設定
    Set ws_処理対象 = ThisWorkbook.Worksheets("データ")
    
    ' 最終行を取得
    lng_最終行 = ws_処理対象.Cells(ws_処理対象.Rows.Count, "A").End(xlUp).Row
    
    ' ヘッダー行をスキップして処理総数を計算
    lng_処理総数 = lng_最終行 - 1
    
    ' 処理総数が0以下の場合は終了
    If lng_処理総数 <= 0 Then
        MsgBox "処理対象のデータがありません。", vbExclamation
        Exit Sub
    End If
    
    ' 開始時間を記録
    dbl_開始時間 = Timer
    
    ' ===== 方法1: UserFormによる進捗表示 =====
    ' UserFormを動的に作成
    Set frm_進捗画面 = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_進捗画面
        .Properties("Caption") = "処理実行中"
        .Properties("Width") = 300
        .Properties("Height") = 120
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' ラベルコントロールを追加（タイトル）
    Dim ctl_タイトル As Object
    Set ctl_タイトル = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_タイトル
        .Left = 10
        .Top = 10
        .Width = 280
        .Height = 20
        .Caption = "データ処理中..."
    End With
    
    ' ラベルコントロールを追加（進捗率）
    Dim ctl_進捗率 As Object
    Set ctl_進捗率 = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_進捗率
        .Left = 10
        .Top = 50
        .Width = 280
        .Height = 20
        .Caption = "0%"
    End With
    
    ' ProgressBarコントロールを追加
    Dim ctl_プログレスバー As Object
    Set ctl_プログレスバー = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_プログレスバー
        .Left = 10
        .Top = 30
        .Width = 280
        .Height = 20
        .BorderStyle = 1 ' Fixed Single
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' ラベルコントロールを追加（残り時間）
    Dim ctl_残り時間 As Object
    Set ctl_残り時間 = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_残り時間
        .Left = 10
        .Top = 70
        .Width = 280
        .Height = 20
        .Caption = "残り時間: 計算中..."
    End With
    
    ' 進捗フォームを表示
    On Error Resume Next
    frm_進捗画面.Designer.Show vbModeless
    DoEvents
    
    ' ===== 方法2: 代替手法としてステータスバーを使用 =====
    Application.DisplayStatusBar = True
    
    ' メインの処理ループ
    For lng_現在行 = 2 To lng_最終行 ' ヘッダー行をスキップ
        ' 進捗率を計算（0〜1の範囲）
        sng_進捗率 = (lng_現在行 - 2) / lng_処理総数
        
        ' ステータスメッセージを構築
        str_ステータス = "処理中... " & Format(sng_進捗率, "0%") & " 完了" & _
                         " (" & lng_現在行 - 2 & "/" & lng_処理総数 & ")"
        
        ' 経過時間と推定残り時間を計算
        dbl_経過時間 = Timer - dbl_開始時間
        If sng_進捗率 > 0 Then
            dbl_推定残り時間 = (dbl_経過時間 / sng_進捗率) - dbl_経過時間
        Else
            dbl_推定残り時間 = 0
        End If
        
        ' 進捗バーを更新（UserForm）
        On Error Resume Next
        If Not frm_進捗画面 Is Nothing Then
            ctl_プログレスバー.Width = 280 * sng_進捗率
            ctl_プログレスバー.BackColor = RGB(0, 176, 240)
            ctl_進捗率.Caption = Format(sng_進捗率, "0%") & " (" & lng_現在行 - 2 & "/" & lng_処理総数 & ")"
            ctl_残り時間.Caption = "残り時間: 約 " & Format(dbl_推定残り時間, "0.0") & " 秒"
            DoEvents
        End If
        
        ' ステータスバーも更新
        Application.StatusBar = str_ステータス & " - 残り時間: 約 " & Format(dbl_推定残り時間, "0.0") & " 秒"
        
        ' ここに実際の処理を記述
        ' サンプルとして少し待機
        Application.Wait Now + TimeSerial(0, 0, 0.05)
        
        ' 例: データの処理
        ws_処理対象.Cells(lng_現在行, "C").Value = "処理済"
        ws_処理対象.Cells(lng_現在行, "D").Value = Now
        
        ' 画面の更新
        DoEvents
    Next lng_現在行
    
    ' 処理完了
    Application.StatusBar = "処理が完了しました。" & lng_処理総数 & " 件のデータを処理しました。"
    
    ' UserFormを閉じる
    On Error Resume Next
    Unload frm_進捗画面
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_進捗画面
    Set frm_進捗画面 = Nothing
    
    ' 完了メッセージを表示
    MsgBox "処理が完了しました。" & vbCrLf & _
           lng_処理総数 & " 件のデータを処理しました。" & vbCrLf & _
           "処理時間: " & Format(Timer - dbl_開始時間, "0.00") & " 秒", vbInformation
    
    ' ステータスバーをクリア
    Application.StatusBar = False
End Sub

' ===== シンプルな進捗表示関数（再利用可能） =====
Public Sub 進捗表示(ByVal lng_現在値 As Long, ByVal lng_最大値 As Long, Optional ByVal str_メッセージ As String = "処理中...")
    Dim sng_進捗率 As Single
    
    ' 進捗率を計算（0〜1の範囲）
    If lng_最大値 > 0 Then
        sng_進捗率 = lng_現在値 / lng_最大値
    Else
        sng_進捗率 = 0
    End If
    
    ' ステータスバーに表示
    Application.StatusBar = str_メッセージ & " " & Format(sng_進捗率, "0%") & _
                           " (" & lng_現在値 & "/" & lng_最大値 & ")"
    
    ' 画面の更新
    DoEvents
End Sub

'''---------------------------------------------------------
' 2. カスタムメッセージボックス・入力フォーム
'''---------------------------------------------------------
Sub カスタムメッセージボックス()
    ' 変数宣言
    Dim str_結果 As String
    
    ' ===== 1. 基本的なカスタムメッセージボックス =====
    str_結果 = カスタム確認メッセージ("処理を実行しますか？", "確認", True)
    
    If str_結果 = "はい" Then
        MsgBox "「はい」が選択されました。処理を実行します。", vbInformation
    ElseIf str_結果 = "いいえ" Then
        MsgBox "「いいえ」が選択されました。処理を中止します。", vbInformation
    ElseIf str_結果 = "キャンセル" Then
        MsgBox "「キャンセル」が選択されました。", vbInformation
    End If
    
    ' ===== 2. カスタム入力フォーム =====
    Dim str_顧客名 As String
    Dim str_メールアドレス As String
    Dim str_コメント As String
    
    ' カスタム入力フォームを表示
    If カスタム入力フォーム(str_顧客名, str_メールアドレス, str_コメント) Then
        ' 入力結果の表示
        MsgBox "入力された情報:" & vbCrLf & _
               "顧客名: " & str_顧客名 & vbCrLf & _
               "メールアドレス: " & str_メールアドレス & vbCrLf & _
               "コメント: " & str_コメント, vbInformation, "入力結果"
    Else
        MsgBox "入力がキャンセルされました。", vbInformation
    End If
End Sub

' カスタム確認メッセージを表示する関数
Public Function カスタム確認メッセージ(str_メッセージ As String, str_タイトル As String, _
                                      Optional bln_キャンセル可能 As Boolean = False) As String
    ' 変数宣言
    Dim frm_メッセージ As Object
    Dim ctl_ラベル As Object
    Dim ctl_はいボタン As Object
    Dim ctl_いいえボタン As Object
    Dim ctl_キャンセルボタン As Object
    Dim lng_フォーム幅 As Long
    Dim lng_フォーム高さ As Long
    Dim lng_ボタン幅 As Long
    Dim lng_ボタン間隔 As Long
    
    ' 初期値の設定
    カスタム確認メッセージ = ""
    lng_フォーム幅 = 300
    lng_フォーム高さ = 140
    lng_ボタン幅 = 80
    lng_ボタン間隔 = 10
    
    ' UserFormを動的に作成
    Set frm_メッセージ = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_メッセージ
        .Properties("Caption") = str_タイトル
        .Properties("Width") = lng_フォーム幅
        .Properties("Height") = lng_フォーム高さ
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' メッセージラベルを追加
    Set ctl_ラベル = frm_メッセージ.Designer.Controls.Add("Forms.Label.1")
    With ctl_ラベル
        .Left = 10
        .Top = 10
        .Width = lng_フォーム幅 - 20
        .Height = 60
        .Caption = str_メッセージ
    End With
    
    ' はいボタンを追加
    Set ctl_はいボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_はいボタン
        .Left = 40
        .Top = lng_フォーム高さ - 40
        .Width = lng_ボタン幅
        .Height = 25
        .Caption = "はい"
        .Name = "ButtonYes"
    End With
    
    ' いいえボタンを追加
    Set ctl_いいえボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_いいえボタン
        .Left = 40 + lng_ボタン幅 + lng_ボタン間隔
        .Top = lng_フォーム高さ - 40
        .Width = lng_ボタン幅
        .Height = 25
        .Caption = "いいえ"
        .Name = "ButtonNo"
    End With
    
    ' キャンセルボタンを追加（オプション）
    If bln_キャンセル可能 Then
        Set ctl_キャンセルボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
        With ctl_キャンセルボタン
            .Left = 40 + (lng_ボタン幅 + lng_ボタン間隔) * 2
            .Top = lng_フォーム高さ - 40
            .Width = lng_ボタン幅
            .Height = 25
            .Caption = "キャンセル"
            .Name = "ButtonCancel"
        End With
    End If
    
    ' ボタンのイベントコードを追加
    With frm_メッセージ.CodeModule
        .InsertLines .CountOfLines + 1, "Private Sub ButtonYes_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""はい"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
    ' 顧客名テキストボックスの追加
    Dim ctl_顧客名テキスト As Object
    Set ctl_顧客名テキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_顧客名テキスト
        .Left = 100
        .Top = 10
        .Width = 180
        .Height = 20
        .Name = "TextName"
        .Text = str_顧客名 ' 既存の値があれば設定
    End With
    
    ' メールアドレスラベルの追加
    Dim ctl_メールラベル As Object
    Set ctl_メールラベル = frm_入力.Designer.Controls.Add("Forms.Label.1")
    With ctl_メールラベル
        .Left = 10
        .Top = 40
        .Width = 80
        .Height = 20
        .Caption = "メール:"
    End With
    
    ' メールアドレステキストボックスの追加
    Dim ctl_メールテキスト As Object
    Set ctl_メールテキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_メールテキスト
        .Left = 100
        .Top = 40
        .Width = 180
        .Height = 20
        .Name = "TextEmail"
        .Text = str_メールアドレス ' 既存の値があれば設定
    End With
    
    ' コメントラベルの追加
    Dim ctl_コメントラベル As Object
    Set ctl_コメントラベル = frm_入力.Designer.Controls.Add("Forms.Label.1")
    With ctl_コメントラベル
        .Left = 10
        .Top = 70
        .Width = 80
        .Height = 20
        .Caption = "コメント:"
    End With
    
    ' コメントテキストボックスの追加（複数行）
    Dim ctl_コメントテキスト As Object
    Set ctl_コメントテキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_コメントテキスト
        .Left = 100
        .Top = 70
        .Width = 180
        .Height = 80
        .MultiLine = True
        .Name = "TextComment"
        .Text = str_コメント ' 既存の値があれば設定
    End With
    
    ' OKボタンの追加
    Dim ctl_OKボタン As Object
    Set ctl_OKボタン = frm_入力.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_OKボタン
        .Left = 70
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "OK"
        .Name = "ButtonOK"
    End With
    
    ' キャンセルボタンの追加
    Dim ctl_キャンセルボタン As Object
    Set ctl_キャンセルボタン = frm_入力.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_キャンセルボタン
        .Left = 160
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "キャンセル"
        .Name = "ButtonCancel"
    End With
    
    ' ボタンのイベントコードを追加
    With frm_入力.CodeModule
        .InsertLines .CountOfLines + 1, "Private Sub ButtonOK_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""OK"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        .InsertLines .CountOfLines + 1, "Private Sub ButtonCancel_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""キャンセル"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
    
    ' フォームを表示
    On Error Resume Next
    frm_入力.Designer.Show
    
    ' 結果を取得
    If frm_入力.Designer.Tag = "OK" Then
        str_顧客名 = frm_入力.Designer.Controls("TextName").Text
        str_メールアドレス = frm_入力.Designer.Controls("TextEmail").Text
        str_コメント = frm_入力.Designer.Controls("TextComment").Text
        カスタム入力フォーム = True
    Else
        カスタム入力フォーム = False
    End If
    
    ' フォームを閉じる
    On Error Resume Next
    Unload frm_入力
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_入力
    Set frm_入力 = Nothing
End Function

'''---------------------------------------------------------
' 3. UserFormの動的コントロール生成と操作
'''---------------------------------------------------------
Sub UserForm動的コントロール生成()
    ' 変数宣言
    Dim frm_動的フォーム As Object
    Dim lng_フォーム幅 As Long
    Dim lng_フォーム高さ As Long
    Dim i As Long
    Dim str_データ配列() As String
    Dim lng_チェック結果 As Long
    
    ' データ配列の設定（サンプル）
    ReDim str_データ配列(1 To 10)
    For i = 1 To 10
        str_データ配列(i) = "項目 " & i
    Next i
    
    ' フォームのサイズ設定
    lng_フォーム幅 = 300
    lng_フォーム高さ = 400
    
    ' UserFormを動的に作成
    Set frm_動的フォーム = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_動的フォーム
        .Properties("Caption") = "動的コントロール生成サンプル"
        .Properties("Width") = lng_フォーム幅
        .Properties("Height") = lng_フォーム高さ
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' ラベルの追加
    Dim ctl_タイトルラベル As Object
    Set ctl_タイトルラベル = frm_動的フォーム.Designer.Controls.Add("Forms.Label.1")
    With ctl_タイトルラベル
        .Left = 10
        .Top = 10
        .Width = lng_フォーム幅 - 20
        .Height = 20
        .Caption = "以下の項目から選択してください:"
        .Font.Bold = True
    End With
    
    ' チェックボックスを動的に生成
    Dim ctl_チェックボックス As Object
    For i = 1 To UBound(str_データ配列)
        Set ctl_チェックボックス = frm_動的フォーム.Designer.Controls.Add("Forms.CheckBox.1")
        With ctl_チェックボックス
            .Left = 20
            .Top = 30 + (i - 1) * 25
            .Width = lng_フォーム幅 - 40
            .Height = 20
            .Caption = str_データ配列(i)
            .Name = "Check" & i
            .Value = False ' 初期値は未チェック
        End With
    Next i
    
    ' 全選択ボタンの追加
    Dim ctl_全選択ボタン As Object
    Set ctl_全選択ボタン = frm_動的フォーム.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_全選択ボタン
        .Left = 20
        .Top = 30 + UBound(str_データ配列) * 25 + 10
        .Width = 120
        .Height = 25
        .Caption = "全て選択"
        .Name = "ButtonSelectAll"
    End With
    
    ' 選択解除ボタンの追加
    Dim ctl_選択解除ボタン As Object
    Set ctl_選択解除ボタン = frm_動的フォーム.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_選択解除ボタン
        .Left = 150
        .Top = 30 + UBound(str_データ配列) * 25 + 10
        .Width = 120
        .Height = 25
        .Caption = "選択解除"
        .Name = "ButtonClearAll"
    End With
    
    ' OKボタンの追加
    Dim ctl_OKボタン As Object
    Set ctl_OKボタン = frm_動的フォーム.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_OKボタン
        .Left = 70
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "OK"
        .Name = "ButtonOK"
    End With
    
    ' キャンセルボタンの追加
    Dim ctl_キャンセルボタン As Object
    Set ctl_キャンセルボタン = frm_動的フォーム.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_キャンセルボタン
        .Left = 160
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "キャンセル"
        .Name = "ButtonCancel"
    End With
    
    ' イベントコードを追加
    With frm_動的フォーム.CodeModule
        ' 全選択ボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonSelectAll_Click()"
        .InsertLines .CountOfLines + 1, "    Dim i As Long"
        .InsertLines .CountOfLines + 1, "    For i = 1 To 10"
        .InsertLines .CountOfLines + 1, "        Me.Controls(""Check"" & i).Value = True"
        .InsertLines .CountOfLines + 1, "    Next i"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        ' 選択解除ボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonClearAll_Click()"
        .InsertLines .CountOfLines + 1, "    Dim i As Long"
        .InsertLines .CountOfLines + 1, "    For i = 1 To 10"
        .InsertLines .CountOfLines + 1, "        Me.Controls(""Check"" & i).Value = False"
        .InsertLines .CountOfLines + 1, "    Next i"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        ' OKボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonOK_Click()"
        .InsertLines .CountOfLines + 1, "    Dim i As Long, lng_結果 As Long"
        .InsertLines .CountOfLines + 1, "    lng_結果 = 0"
        .InsertLines .CountOfLines + 1, "    For i = 1 To 10"
        .InsertLines .CountOfLines + 1, "        If Me.Controls(""Check"" & i).Value Then"
        .InsertLines .CountOfLines + 1, "            lng_結果 = lng_結果 + 1"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "    Next i"
        .InsertLines .CountOfLines + 1, "    Tag = lng_結果"
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        ' キャンセルボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonCancel_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""キャンセル"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
    
    ' フォームを表示
    On Error Resume Next
    frm_動的フォーム.Designer.Show
    
    ' 結果を取得
    If frm_動的フォーム.Designer.Tag <> "キャンセル" Then
        lng_チェック結果 = CLng(frm_動的フォーム.Designer.Tag)
        MsgBox "選択された項目数: " & lng_チェック結果, vbInformation
    Else
        MsgBox "キャンセルされました。", vbInformation
    End If
    
    ' フォームを閉じる
    On Error Resume Next
    Unload frm_動的フォーム
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_動的フォーム
    Set frm_動的フォーム = Nothing
End Sub

'''---------------------------------------------------------
' 4. セル上でのコンテキストメニュー実装
'''---------------------------------------------------------
Sub コンテキストメニュー実装()
    ' 変数宣言
    Dim ws_対象 As Worksheet
    Dim ContextMenuExists As Boolean
    
    ' 対象ワークシートの設定
    Set ws_対象 = ThisWorkbook.Worksheets("Sheet1")
    
    ' 既存のカスタムメニューがあるか確認
    On Error Resume Next
    ContextMenuExists = Not (Application.CommandBars("セル右クリックメニュー") Is Nothing)
    On Error GoTo 0
    
    ' 既存のメニューがあれば削除
    If ContextMenuExists Then
        Application.CommandBars("セル右クリックメニュー").Delete
    End If
    
    ' 新しいコンテキストメニューを作成
    Dim cb_メニュー As CommandBar
    Dim ctrl_メニュー項目 As CommandBarButton
    
    ' コマンドバーの作成
    Set cb_メニュー = Application.CommandBars.Add(Name:="セル右クリックメニュー", Position:=msoBarPopup, MenuBar:=False, Temporary:=True)
    
    ' メニュー項目1: データクリア
    Set ctrl_メニュー項目 = cb_メニュー.Controls.Add(Type:=msoControlButton)
    With ctrl_メニュー項目
        .Caption = "データをクリア"
        .OnAction = "データクリア_コンテキストメニュー"
        .FaceId = 288  ' アイコンID
    End With
    
    ' メニュー項目2: 書式だけクリア
    Set ctrl_メニュー項目 = cb_メニュー.Controls.Add(Type:=msoControlButton)
    With ctrl_メニュー項目
        .Caption = "書式だけクリア"
        .OnAction = "書式クリア_コンテキストメニュー"
        .FaceId = 37   ' アイコンID
    End With
    
    ' セパレーターの追加
    Set ctrl_メニュー項目 = cb_メニュー.Controls.Add(Type:=msoControlButton)
    ctrl_メニュー項目.BeginGroup = True
    
    ' メニュー項目3: サブメニュー付き項目
    With ctrl_メニュー項目
        .Caption = "データ変換"
        .OnAction = ""  ' サブメニューなのでアクションは不要
        .FaceId = 586   ' アイコンID
    End With
    
    ' サブメニュー1: 大文字に変換
    Dim ctrl_サブメニュー As CommandBarButton
    Set ctrl_サブメニュー = cb_メニュー.Controls.Add(Type:=msoControlButton)
    With ctrl_サブメニュー
        .Caption = "　大文字に変換"
        .OnAction = "大文字変換_コンテキストメニュー"
        .FaceId = 104  ' アイコンID
        .BeginGroup = True
    End With
    
    ' サブメニュー2: 小文字に変換
    Set ctrl_サブメニュー = cb_メニュー.Controls.Add(Type:=msoControlButton)
    With ctrl_サブメニュー
        .Caption = "　小文字に変換"
        .OnAction = "小文字変換_コンテキストメニュー"
        .FaceId = 103  ' アイコンID
    End With
    
    ' セパレーターの追加
    Set ctrl_メニュー項目 = cb_メニュー.Controls.Add(Type:=msoControlButton)
    ctrl_メニュー項目.BeginGroup = True
    
    ' メニュー項目4: カスタム書式設定
    With ctrl_メニュー項目
        .Caption = "カスタム書式設定"
        .OnAction = "カスタム書式_コンテキストメニュー"
        .FaceId = 186  ' アイコンID
    End With
    
    ' メニューをセルのコンテキストメニューに追加
    Application.CommandBars("Cell").Controls.Add Type:=msoControlButton, _
                                              before:=1, temporary:=True
    
    ' メッセージ表示
    MsgBox "カスタムコンテキストメニューが作成されました。" & vbCrLf & _
           "対象シートのセル上で右クリックして確認してください。", vbInformation
End Sub

' コンテキストメニューの処理関数
Public Sub データクリア_コンテキストメニュー()
    On Error Resume Next
    Selection.ClearContents
End Sub

Public Sub 書式クリア_コンテキストメニュー()
    On Error Resume Next
    Selection.ClearFormats
End Sub

Public Sub 大文字変換_コンテキストメニュー()
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Dim rng_選択範囲 As Range
        Set rng_選択範囲 = Selection
        
        Dim cel As Range
        For Each cel In rng_選択範囲
            If Not IsEmpty(cel.Value) Then
                cel.Value = UCase(cel.Value)
            End If
        Next cel
    End If
End Sub

Public Sub 小文字変換_コンテキストメニュー()
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Dim rng_選択範囲 As Range
        Set rng_選択範囲 = Selection
        
        Dim cel As Range
        For Each cel In rng_選択範囲
            If Not IsEmpty(cel.Value) Then
                cel.Value = LCase(cel.Value)
            End If
        Next cel
    End If
End Sub

Public Sub カスタム書式_コンテキストメニュー()
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        ' カスタム書式設定ダイアログの代わりに単純な書式を適用
        Selection.NumberFormat = "yyyy/mm/dd"
        Selection.Font.Bold = True
        Selection.Interior.Color = RGB(255, 255, 200)  ' 薄い黄色
    End If
End Sub

'''---------------------------------------------------------
' 5. 入力値の検証とフィードバック提供
'''---------------------------------------------------------
Sub 入力値検証フィードバック()
    ' 変数宣言
    Dim str_入力値 As String
    Dim bln_有効 As Boolean
    Dim str_フィードバック As String
    
    ' 入力値の取得
    str_入力値 = InputBox("メールアドレスを入力してください:", "メールアドレス入力")
    
    ' キャンセルされた場合は終了
    If str_入力値 = "" Then Exit Sub
    
    ' メールアドレスの検証
    bln_有効 = メールアドレス検証(str_入力値, str_フィードバック)
    
    ' 結果の表示
    If bln_有効 Then
        MsgBox "入力されたメールアドレスは有効です: " & str_入力値, vbInformation
    Else
        MsgBox "無効なメールアドレスです: " & str_入力値 & vbCrLf & _
               "理由: " & str_フィードバック, vbExclamation
    End If
    
    ' より高度な入力検証のためのフォームを表示
    Call 高度な入力検証フォーム
End Sub

' メールアドレスを検証する関数
Public Function メールアドレス検証(str_メールアドレス As String, ByRef str_エラーメッセージ As String) As Boolean
    ' 初期設定
    メールアドレス検証 = False
    str_エラーメッセージ = ""
    
    ' 空文字チェック
    If Trim(str_メールアドレス) = "" Then
        str_エラーメッセージ = "メールアドレスが入力されていません。"
        Exit Function
    End If
    
    ' @記号の存在チェック
    If InStr(str_メールアドレス, "@") = 0 Then
        str_エラーメッセージ = "@記号が含まれていません。"
        Exit Function
    End If
    
    ' @の位置が適切かチェック
    If InStr(str_メールアドレス, "@") = 1 Then
        str_エラーメッセージ = "ローカル部が存在しません。"
        Exit Function
    End If
    
    ' 複数の@記号チェック
    If Len(str_メールアドレス) - Len(Replace(str_メールアドレス, "@", "")) > 1 Then
        str_エラーメッセージ = "@記号が複数含まれています。"
        Exit Function
    End If
    
    ' ドメイン部分の存在チェック
    Dim arr_部分 As Variant
    arr_部分 = Split(str_メールアドレス, "@")
    
    If arr_部分(1) = "" Then
        str_エラーメッセージ = "ドメイン部が存在しません。"
        Exit Function
    End If
    
    ' ドメイン部にドットが含まれているかチェック
    If InStr(arr_部分(1), ".") = 0 Then
        str_エラーメッセージ = "ドメイン部にドット（.）が含まれていません。"
        Exit Function
    End If
    
    ' ドメイン名が適切かチェック
    Dim arr_ドメイン As Variant
    arr_ドメイン = Split(arr_部分(1), ".")
    
    If arr_ドメイン(UBound(arr_ドメイン)) = "" Then
        str_エラーメッセージ = "トップレベルドメインが不正です。"
        Exit Function
    End If
    
    ' 禁則文字のチェック
    Dim str_禁則文字 As String
    str_禁則文字 = " ,;:!#$%^&*()=+<>{}[]|\/"
    
    Dim i As Long
    For i = 1 To Len(str_禁則文字)
        If InStr(str_メールアドレス, Mid(str_禁則文字, i, 1)) > 0 Then
            str_エラーメッセージ = "禁止文字（" & Mid(str_禁則文字, i, 1) & "）が含まれています。"
            Exit Function
        End If
    Next i
    
    ' すべての検証に通過
    メールアドレス検証 = True
End Function

' 高度な入力検証フォームを表示する関数
Public Sub 高度な入力検証フォーム()
    ' 変数宣言
    Dim frm_検証 As Object
    Dim lng_フォーム幅 As Long
    Dim lng_フォーム高さ As Long
    
    ' フォームのサイズ設定
    lng_フォーム幅 = 380
    lng_フォーム高さ = 320
    
    ' UserFormを動的に作成
    Set frm_検証 = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_検証
        .Properties("Caption") = "入力値検証サンプル"
        .Properties("Width") = lng_フォーム幅
        .Properties("Height") = lng_フォーム高さ
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' フォームにコントロールを追加
    ' --- 名前フィールド ---
    Dim ctl_名前ラベル As Object
    Set ctl_名前ラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_名前ラベル
        .Left = 10
        .Top = 10
        .Width = 120
        .Height = 20
        .Caption = "氏名 (必須):"
    End With
    
    Dim ctl_名前テキスト As Object
    Set ctl_名前テキスト = frm_検証.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_名前テキスト
        .Left = 140
        .Top = 10
        .Width = 220
        .Height = 20
        .Name = "TextName"
    End With
    
    ' --- 電話番号フィールド ---
    Dim ctl_電話ラベル As Object
    Set ctl_電話ラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_電話ラベル
        .Left = 10
        .Top = 40
        .Width = 120
        .Height = 20
        .Caption = "電話番号:"
    End With
    
    Dim ctl_電話テキスト As Object
    Set ctl_電話テキスト = frm_検証.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_電話テキスト
        .Left = 140
        .Top = 40
        .Width = 220
        .Height = 20
        .Name = "TextPhone"
    End With
    
    ' --- 郵便番号フィールド ---
    Dim ctl_郵便ラベル As Object
    Set ctl_郵便ラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_郵便ラベル
        .Left = 10
        .Top = 70
        .Width = 120
        .Height = 20
        .Caption = "郵便番号:"
    End With
    
    Dim ctl_郵便テキスト As Object
    Set ctl_郵便テキスト = frm_検証.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_郵便テキスト
        .Left = 140
        .Top = 70
        .Width = 220
        .Height = 20
        .Name = "TextZip"
    End With
    
    ' --- 生年月日フィールド ---
    Dim ctl_生年月日ラベル As Object
    Set ctl_生年月日ラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_生年月日ラベル
        .Left = 10
        .Top = 100
        .Width = 120
        .Height = 20
        .Caption = "生年月日:"
    End With
    
    Dim ctl_生年月日テキスト As Object
    Set ctl_生年月日テキスト = frm_検証.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_生年月日テキスト
        .Left = 140
        .Top = 100
        .Width = 220
        .Height = 20
        .Name = "TextBirthday"
    End With
    
    ' --- メールアドレスフィールド ---
    Dim ctl_メールラベル As Object
    Set ctl_メールラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_メールラベル
        .Left = 10
        .Top = 130
        .Width = 120
        .Height = 20
        .Caption = "メールアドレス (必須):"
    End With
    
    Dim ctl_メールテキスト As Object
    Set ctl_メールテキスト = frm_検証.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_メールテキスト
        .Left = 140
        .Top = 130
        .Width = 220
        .Height = 20
        .Name = "TextEmail"
    End With
    
    ' --- フィードバックラベル ---
    Dim ctl_フィードバックラベル As Object
    Set ctl_フィードバックラベル = frm_検証.Designer.Controls.Add("Forms.Label.1")
    With ctl_フィードバックラベル
        .Left = 10
        .Top = 170
        .Width = 350
        .Height = 80
        .Name = "LabelFeedback"
        .BorderStyle = 1 ' 枠線あり
        .BackColor = RGB(255, 255, 200) ' 薄い黄色
        .Caption = "入力内容を確認してください。必須フィールドは必ず入力してください。"
    End With
    
    ' --- 検証ボタン ---
    Dim ctl_検証ボタン As Object
    Set ctl_検証ボタン = frm_検証.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_検証ボタン
        .Left = 100
        .Top = 260
        .Width = 80
        .Height = 25
        .Caption = "検証"
        .Name = "ButtonValidate"
    End With
    
    ' --- OKボタン ---
    Dim ctl_OKボタン As Object
    Set ctl_OKボタン = frm_検証.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_OKボタン
        .Left = 190
        .Top = 260
        .Width = 80
        .Height = 25
        .Caption = "OK"
        .Name = "ButtonOK"
        .Enabled = False ' 初期状態では無効
    End With
    
    ' --- キャンセルボタン ---
    Dim ctl_キャンセルボタン As Object
    Set ctl_キャンセルボタン = frm_検証.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_キャンセルボタン
        .Left = 280
        .Top = 260
        .Width = 80
        .Height = 25
        .Caption = "キャンセル"
        .Name = "ButtonCancel"
    End With
    
    ' イベントコードを追加
    With frm_検証.CodeModule
        ' 検証ボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonValidate_Click()"
        .InsertLines .CountOfLines + 1, "    Dim bln_有効 As Boolean"
        .InsertLines .CountOfLines + 1, "    Dim str_エラーメッセージ As String"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' 必須フィールドのチェック"
        .InsertLines .CountOfLines + 1, "    If Trim(TextName.Text) = """" Then"
        .InsertLines .CountOfLines + 1, "        LabelFeedback.Caption = ""氏名は必須項目です。入力してください。"""
        .InsertLines .CountOfLines + 1, "        TextName.SetFocus"
        .InsertLines .CountOfLines + 1, "        Exit Sub"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    If Trim(TextEmail.Text) = """" Then"
        .InsertLines .CountOfLines + 1, "        LabelFeedback.Caption = ""メールアドレスは必須項目です。入力してください。"""
        .InsertLines .CountOfLines + 1, "        TextEmail.SetFocus"
        .InsertLines .CountOfLines + 1, "        Exit Sub"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' メールアドレスの検証"
        .InsertLines .CountOfLines + 1, "    If InStr(TextEmail.Text, ""@"") = 0 Or InStr(TextEmail.Text, ""."") = 0 Then"
        .InsertLines .CountOfLines + 1, "        LabelFeedback.Caption = ""メールアドレスの形式が正しくありません。"""
        .InsertLines .CountOfLines + 1, "        TextEmail.SetFocus"
        .InsertLines .CountOfLines + 1, "        Exit Sub"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' 電話番号の検証（入力されている場合）"
        .InsertLines .CountOfLines + 1, "    If Trim(TextPhone.Text) <> """" Then"
        .InsertLines .CountOfLines + 1, "        Dim str_電話 As String"
        .InsertLines .CountOfLines + 1, "        str_電話 = Replace(Replace(Replace(TextPhone.Text, ""-"", """"), ""("", """"), "")"", """")"
        .InsertLines .CountOfLines + 1, "        If Not IsNumeric(str_電話) Then"
        .InsertLines .CountOfLines + 1, "            LabelFeedback.Caption = ""電話番号は数字と一部の記号のみ使用できます。"""
        .InsertLines .CountOfLines + 1, "            TextPhone.SetFocus"
        .InsertLines .CountOfLines + 1, "            Exit Sub"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' 郵便番号の検証（入力されている場合）"
        .InsertLines .CountOfLines + 1, "    If Trim(TextZip.Text) <> """" Then"
        .InsertLines .CountOfLines + 1, "        Dim str_郵便 As String"
        .InsertLines .CountOfLines + 1, "        str_郵便 = Replace(TextZip.Text, ""-"", """")"
        .InsertLines .CountOfLines + 1, "        If Not IsNumeric(str_郵便) Then"
        .InsertLines .CountOfLines + 1, "            LabelFeedback.Caption = ""郵便番号は数字とハイフンのみ使用できます。"""
        .InsertLines .CountOfLines + 1, "            TextZip.SetFocus"
        .InsertLines .CountOfLines + 1, "            Exit Sub"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' 生年月日の検証（入力されている場合）"
        .InsertLines .CountOfLines + 1, "    If Trim(TextBirthday.Text) <> """" Then"
        .InsertLines .CountOfLines + 1, "        Dim dt_誕生日 As Date"
        .InsertLines .CountOfLines + 1, "        On Error Resume Next"
        .InsertLines .CountOfLines + 1, "        dt_誕生日 = CDate(TextBirthday.Text)"
        .InsertLines .CountOfLines + 1, "        If Err.Number <> 0 Then"
        .InsertLines .CountOfLines + 1, "            LabelFeedback.Caption = ""生年月日の形式が正しくありません。YYYY/MM/DD形式で入力してください。"""
        .InsertLines .CountOfLines + 1, "            Err.Clear"
        .InsertLines .CountOfLines + 1, "            TextBirthday.SetFocus"
        .InsertLines .CountOfLines + 1, "            Exit Sub"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "        On Error GoTo 0"
        .InsertLines .CountOfLines + 1, "        "
        .InsertLines .CountOfLines + 1, "        ' 未来の日付でないことを確認"
        .InsertLines .CountOfLines + 1, "        If dt_誕生日 > Date Then"
        .InsertLines .CountOfLines + 1, "            LabelFeedback.Caption = ""生年月日が未来の日付になっています。"""
        .InsertLines .CountOfLines + 1, "            TextBirthday.SetFocus"
        .InsertLines .CountOfLines + 1, "            Exit Sub"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    ' すべての検証に通過した場合"
        .InsertLines .CountOfLines + 1, "    LabelFeedback.Caption = ""すべての入力が有効です。OKボタンを押して続行できます。"""
        .InsertLines .CountOfLines + 1, "    LabelFeedback.BackColor = RGB(200, 255, 200) ' 薄い緑色"
        .InsertLines .CountOfLines + 1, "    ButtonOK.Enabled = True"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        ' OKボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonOK_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""OK"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        ' キャンセルボタンのイベント
        .InsertLines .CountOfLines + 1, "Private Sub ButtonCancel_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""キャンセル"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
    
    ' フォームを表示
    On Error Resume Next
    frm_検証.Designer.Show
    
    ' 結果を取得
    If frm_検証.Designer.Tag = "OK" Then
        MsgBox "入力データが検証され、有効と判断されました。", vbInformation
    Else
        MsgBox "入力がキャンセルされました。", vbInformation
    End If
    
    ' フォームを閉じる
    On Error Resume Next
    Unload frm_検証
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_検証
    Set frm_検証 = Nothing
End Sub
        '''---------------------------------------------------------
' 1. 進捗バーの実装と進捗状況表示
'''---------------------------------------------------------
Sub 進捗バー実装()
    ' 変数宣言
    Dim frm_進捗画面 As Object
    Dim ws_処理対象 As Worksheet
    Dim lng_最終行 As Long
    Dim lng_現在行 As Long
    Dim lng_処理総数 As Long
    Dim sng_進捗率 As Single
    Dim str_ステータス As String
    Dim dbl_開始時間 As Double
    Dim dbl_経過時間 As Double
    Dim dbl_推定残り時間 As Double
    
    ' 処理対象のワークシートを設定
    Set ws_処理対象 = ThisWorkbook.Worksheets("データ")
    
    ' 最終行を取得
    lng_最終行 = ws_処理対象.Cells(ws_処理対象.Rows.Count, "A").End(xlUp).Row
    
    ' ヘッダー行をスキップして処理総数を計算
    lng_処理総数 = lng_最終行 - 1
    
    ' 処理総数が0以下の場合は終了
    If lng_処理総数 <= 0 Then
        MsgBox "処理対象のデータがありません。", vbExclamation
        Exit Sub
    End If
    
    ' 開始時間を記録
    dbl_開始時間 = Timer
    
    ' ===== 方法1: UserFormによる進捗表示 =====
    ' UserFormを動的に作成
    Set frm_進捗画面 = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_進捗画面
        .Properties("Caption") = "処理実行中"
        .Properties("Width") = 300
        .Properties("Height") = 120
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' ラベルコントロールを追加（タイトル）
    Dim ctl_タイトル As Object
    Set ctl_タイトル = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_タイトル
        .Left = 10
        .Top = 10
        .Width = 280
        .Height = 20
        .Caption = "データ処理中..."
    End With
    
    ' ラベルコントロールを追加（進捗率）
    Dim ctl_進捗率 As Object
    Set ctl_進捗率 = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_進捗率
        .Left = 10
        .Top = 50
        .Width = 280
        .Height = 20
        .Caption = "0%"
    End With
    
    ' ProgressBarコントロールを追加
    Dim ctl_プログレスバー As Object
    Set ctl_プログレスバー = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_プログレスバー
        .Left = 10
        .Top = 30
        .Width = 280
        .Height = 20
        .BorderStyle = 1 ' Fixed Single
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' ラベルコントロールを追加（残り時間）
    Dim ctl_残り時間 As Object
    Set ctl_残り時間 = frm_進捗画面.Designer.Controls.Add("Forms.Label.1")
    With ctl_残り時間
        .Left = 10
        .Top = 70
        .Width = 280
        .Height = 20
        .Caption = "残り時間: 計算中..."
    End With
    
    ' 進捗フォームを表示
    On Error Resume Next
    frm_進捗画面.Designer.Show vbModeless
    DoEvents
    
    ' ===== 方法2: 代替手法としてステータスバーを使用 =====
    Application.DisplayStatusBar = True
    
    ' メインの処理ループ
    For lng_現在行 = 2 To lng_最終行 ' ヘッダー行をスキップ
        ' 進捗率を計算（0〜1の範囲）
        sng_進捗率 = (lng_現在行 - 2) / lng_処理総数
        
        ' ステータスメッセージを構築
        str_ステータス = "処理中... " & Format(sng_進捗率, "0%") & " 完了" & _
                         " (" & lng_現在行 - 2 & "/" & lng_処理総数 & ")"
        
        ' 経過時間と推定残り時間を計算
        dbl_経過時間 = Timer - dbl_開始時間
        If sng_進捗率 > 0 Then
            dbl_推定残り時間 = (dbl_経過時間 / sng_進捗率) - dbl_経過時間
        Else
            dbl_推定残り時間 = 0
        End If
        
        ' 進捗バーを更新（UserForm）
        On Error Resume Next
        If Not frm_進捗画面 Is Nothing Then
            ctl_プログレスバー.Width = 280 * sng_進捗率
            ctl_プログレスバー.BackColor = RGB(0, 176, 240)
            ctl_進捗率.Caption = Format(sng_進捗率, "0%") & " (" & lng_現在行 - 2 & "/" & lng_処理総数 & ")"
            ctl_残り時間.Caption = "残り時間: 約 " & Format(dbl_推定残り時間, "0.0") & " 秒"
            DoEvents
        End If
        
        ' ステータスバーも更新
        Application.StatusBar = str_ステータス & " - 残り時間: 約 " & Format(dbl_推定残り時間, "0.0") & " 秒"
        
        ' ここに実際の処理を記述
        ' サンプルとして少し待機
        Application.Wait Now + TimeSerial(0, 0, 0.05)
        
        ' 例: データの処理
        ws_処理対象.Cells(lng_現在行, "C").Value = "処理済"
        ws_処理対象.Cells(lng_現在行, "D").Value = Now
        
        ' 画面の更新
        DoEvents
    Next lng_現在行
    
    ' 処理完了
    Application.StatusBar = "処理が完了しました。" & lng_処理総数 & " 件のデータを処理しました。"
    
    ' UserFormを閉じる
    On Error Resume Next
    Unload frm_進捗画面
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_進捗画面
    Set frm_進捗画面 = Nothing
    
    ' 完了メッセージを表示
    MsgBox "処理が完了しました。" & vbCrLf & _
           lng_処理総数 & " 件のデータを処理しました。" & vbCrLf & _
           "処理時間: " & Format(Timer - dbl_開始時間, "0.00") & " 秒", vbInformation
    
    ' ステータスバーをクリア
    Application.StatusBar = False
End Sub

' ===== シンプルな進捗表示関数（再利用可能） =====
Public Sub 進捗表示(ByVal lng_現在値 As Long, ByVal lng_最大値 As Long, Optional ByVal str_メッセージ As String = "処理中...")
    Dim sng_進捗率 As Single
    
    ' 進捗率を計算（0〜1の範囲）
    If lng_最大値 > 0 Then
        sng_進捗率 = lng_現在値 / lng_最大値
    Else
        sng_進捗率 = 0
    End If
    
    ' ステータスバーに表示
    Application.StatusBar = str_メッセージ & " " & Format(sng_進捗率, "0%") & _
                           " (" & lng_現在値 & "/" & lng_最大値 & ")"
    
    ' 画面の更新
    DoEvents
End Sub

'''---------------------------------------------------------
' 2. カスタムメッセージボックス・入力フォーム
'''---------------------------------------------------------
Sub カスタムメッセージボックス()
    ' 変数宣言
    Dim str_結果 As String
    
    ' ===== 1. 基本的なカスタムメッセージボックス =====
    str_結果 = カスタム確認メッセージ("処理を実行しますか？", "確認", True)
    
    If str_結果 = "はい" Then
        MsgBox "「はい」が選択されました。処理を実行します。", vbInformation
    ElseIf str_結果 = "いいえ" Then
        MsgBox "「いいえ」が選択されました。処理を中止します。", vbInformation
    ElseIf str_結果 = "キャンセル" Then
        MsgBox "「キャンセル」が選択されました。", vbInformation
    End If
    
    ' ===== 2. カスタム入力フォーム =====
    Dim str_顧客名 As String
    Dim str_メールアドレス As String
    Dim str_コメント As String
    
    ' カスタム入力フォームを表示
    If カスタム入力フォーム(str_顧客名, str_メールアドレス, str_コメント) Then
        ' 入力結果の表示
        MsgBox "入力された情報:" & vbCrLf & _
               "顧客名: " & str_顧客名 & vbCrLf & _
               "メールアドレス: " & str_メールアドレス & vbCrLf & _
               "コメント: " & str_コメント, vbInformation, "入力結果"
    Else
        MsgBox "入力がキャンセルされました。", vbInformation
    End If
End Sub

' カスタム確認メッセージを表示する関数
Public Function カスタム確認メッセージ(str_メッセージ As String, str_タイトル As String, _
                                      Optional bln_キャンセル可能 As Boolean = False) As String
    ' 変数宣言
    Dim frm_メッセージ As Object
    Dim ctl_ラベル As Object
    Dim ctl_はいボタン As Object
    Dim ctl_いいえボタン As Object
    Dim ctl_キャンセルボタン As Object
    Dim lng_フォーム幅 As Long
    Dim lng_フォーム高さ As Long
    Dim lng_ボタン幅 As Long
    Dim lng_ボタン間隔 As Long
    
    ' 初期値の設定
    カスタム確認メッセージ = ""
    lng_フォーム幅 = 300
    lng_フォーム高さ = 140
    lng_ボタン幅 = 80
    lng_ボタン間隔 = 10
    
    ' UserFormを動的に作成
    Set frm_メッセージ = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_メッセージ
        .Properties("Caption") = str_タイトル
        .Properties("Width") = lng_フォーム幅
        .Properties("Height") = lng_フォーム高さ
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' メッセージラベルを追加
    Set ctl_ラベル = frm_メッセージ.Designer.Controls.Add("Forms.Label.1")
    With ctl_ラベル
        .Left = 10
        .Top = 10
        .Width = lng_フォーム幅 - 20
        .Height = 60
        .Caption = str_メッセージ
    End With
    
    ' はいボタンを追加
    Set ctl_はいボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_はいボタン
        .Left = 40
        .Top = lng_フォーム高さ - 40
        .Width = lng_ボタン幅
        .Height = 25
        .Caption = "はい"
        .Name = "ButtonYes"
    End With
    
    ' いいえボタンを追加
    Set ctl_いいえボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_いいえボタン
        .Left = 40 + lng_ボタン幅 + lng_ボタン間隔
        .Top = lng_フォーム高さ - 40
        .Width = lng_ボタン幅
        .Height = 25
        .Caption = "いいえ"
        .Name = "ButtonNo"
    End With
    
    ' キャンセルボタンを追加（オプション）
    If bln_キャンセル可能 Then
        Set ctl_キャンセルボタン = frm_メッセージ.Designer.Controls.Add("Forms.CommandButton.1")
        With ctl_キャンセルボタン
            .Left = 40 + (lng_ボタン幅 + lng_ボタン間隔) * 2
            .Top = lng_フォーム高さ - 40
            .Width = lng_ボタン幅
            .Height = 25
            .Caption = "キャンセル"
            .Name = "ButtonCancel"
        End With
    End If
    
    ' ボタンのイベントコードを追加
    With frm_メッセージ.CodeModule
        .InsertLines .CountOfLines + 1, "Private Sub ButtonYes_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""はい"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        .InsertLines .CountOfLines + 1, "Private Sub ButtonNo_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""いいえ"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        If bln_キャンセル可能 Then
            .InsertLines .CountOfLines + 1, "Private Sub ButtonCancel_Click()"
            .InsertLines .CountOfLines + 1, "    Tag = ""キャンセル"""
            .InsertLines .CountOfLines + 1, "    Hide"
            .InsertLines .CountOfLines + 1, "End Sub"
        End If
    End With
    
    ' フォームを表示
    On Error Resume Next
    frm_メッセージ.Designer.Show
    
    ' 結果を取得
    カスタム確認メッセージ = frm_メッセージ.Designer.Tag
    
    ' フォームを閉じる
    On Error Resume Next
    Unload frm_メッセージ
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_メッセージ
    Set frm_メッセージ = Nothing
End Function

' カスタム入力フォームを表示する関数
Public Function カスタム入力フォーム(ByRef str_顧客名 As String, ByRef str_メールアドレス As String, _
                                   ByRef str_コメント As String) As Boolean
    ' 変数宣言
    Dim frm_入力 As Object
    Dim lng_フォーム幅 As Long
    Dim lng_フォーム高さ As Long
    
    ' 初期値の設定
    カスタム入力フォーム = False
    lng_フォーム幅 = 300
    lng_フォーム高さ = 220
    
    ' UserFormを動的に作成
    Set frm_入力 = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3=vbext_ct_MSForm
    
    ' フォームのプロパティを設定
    With frm_入力
        .Properties("Caption") = "顧客情報入力"
        .Properties("Width") = lng_フォーム幅
        .Properties("Height") = lng_フォーム高さ
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' 顧客名ラベルの追加
    Dim ctl_顧客名ラベル As Object
    Set ctl_顧客名ラベル = frm_入力.Designer.Controls.Add("Forms.Label.1")
    With ctl_顧客名ラベル
        .Left = 10
        .Top = 10
        .Width = 80
        .Height = 20
        .Caption = "顧客名:"
    End With
    
    ' 顧客名テキストボックスの追加
    Dim ctl_顧客名テキスト As Object
    Set ctl_顧客名テキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_顧客名テキスト
        .Left = 100
        .Top = 10
        .Width = 180
        .Height = 20
        .Name = "TextName"
        .Text = str_顧客名 ' 既存の値があれば設定
    End With
    
    ' メールアドレスラベルの追加
    Dim ctl_メールラベル As Object
    Set ctl_メールラベル = frm_入力.Designer.Controls.Add("Forms.Label.1")
    With ctl_メールラベル
        .Left = 10
        .Top = 40
        .Width = 80
        .Height = 20
        .Caption = "メール:"
    End With
    
    ' メールアドレステキストボックスの追加
    Dim ctl_メールテキスト As Object
    Set ctl_メールテキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_メールテキスト
        .Left = 100
        .Top = 40
        .Width = 180
        .Height = 20
        .Name = "TextEmail"
        .Text = str_メールアドレス ' 既存の値があれば設定
    End With
    
    ' コメントラベルの追加
    Dim ctl_コメントラベル As Object
    Set ctl_コメントラベル = frm_入力.Designer.Controls.Add("Forms.Label.1")
    With ctl_コメントラベル
        .Left = 10
        .Top = 70
        .Width = 80
        .Height = 20
        .Caption = "コメント:"
    End With
    
    ' コメントテキストボックスの追加（複数行）
    Dim ctl_コメントテキスト As Object
    Set ctl_コメントテキスト = frm_入力.Designer.Controls.Add("Forms.TextBox.1")
    With ctl_コメントテキスト
        .Left = 100
        .Top = 70
        .Width = 180
        .Height = 80
        .Name = "TextComment"
        .Text = str_コメント ' 既存の値があれば設定
        .MultiLine = True ' 複数行入力可能に設定
    End With
    
    ' OKボタンの追加
    Dim ctl_OKボタン As Object
    Set ctl_OKボタン = frm_入力.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_OKボタン
        .Left = 60
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "OK"
        .Name = "ButtonOK"
    End With
    
    ' キャンセルボタンの追加
    Dim ctl_キャンセルボタン As Object
    Set ctl_キャンセルボタン = frm_入力.Designer.Controls.Add("Forms.CommandButton.1")
    With ctl_キャンセルボタン
        .Left = 160
        .Top = lng_フォーム高さ - 40
        .Width = 80
        .Height = 25
        .Caption = "キャンセル"
        .Name = "ButtonCancel"
    End With
    
    ' ボタンのイベントコードを追加
    With frm_入力.CodeModule
        .InsertLines .CountOfLines + 1, "Private Sub ButtonOK_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""OK"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
        
        .InsertLines .CountOfLines + 1, "Private Sub ButtonCancel_Click()"
        .InsertLines .CountOfLines + 1, "    Tag = ""キャンセル"""
        .InsertLines .CountOfLines + 1, "    Hide"
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
    
    ' フォームを表示
    On Error Resume Next
    frm_入力.Designer.Show
    
    ' 結果を取得
    If frm_入力.Designer.Tag = "OK" Then
        ' OKボタンが押された場合、入力値を取得
        str_顧客名 = frm_入力.Designer.Controls("TextName").Text
        str_メールアドレス = frm_入力.Designer.Controls("TextEmail").Text
        str_コメント = frm_入力.Designer.Controls("TextComment").Text
        カスタム入力フォーム = True
    End If
    
    ' フォームを閉じる
    On Error Resume Next
    Unload frm_入力
    
    ' 後片付け
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frm_入力
    Set frm_入力 = Nothing
End Function
