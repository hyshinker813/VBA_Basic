Sub ハイフン前の数字を1増やす_改()

    '--- 設定 ---
    Dim str_対象列 As String
    Dim lng_開始行 As Long
    Dim lng_終了行 As Long
    
    '--- 正規表現オブジェクトの宣言 ---
    Dim obj_正規表現 As Object
    Set obj_正規表現 = CreateObject("VBScript.RegExp")

    '--- 変数の初期設定 (ここを必要に応じて変更してください) ---
    str_対象列 = "A" ' 例: A列を対象とする
    lng_開始行 = 1       ' 例: 1行目から処理を開始する
    lng_終了行 = 0         ' 例: 最終行まで処理する場合 (0を設定)
    
    ' 検索する正規表現パターン: 「数字の並び-数字の並び」
    obj_正規表現.Pattern = "(\d+)-(\d+)"
    obj_正規表現.Global = True ' セル内のすべての一致箇所を処理する

    '--- 最終行の特定 (終了行が0の場合) ---
    If lng_終了行 = 0 Then
        lng_終了行 = Cells(Rows.Count, str_対象列).End(xlUp).Row
    End If

    '--- 列内のセルをループして置換 ---
    Dim lng_行番号 As Long
    Dim str_セル内の元のテキスト As String
    Dim obj_見つかった全ての候補 As Object ' MatchCollection
    Dim obj_見つかった一つの候補 As Object   ' Match
    Dim lng_ハイフン前の数字 As Long
    Dim str_ハイフン後の数字 As String
    Dim str_置き換える文字列 As String

    For lng_行番号 = lng_開始行 To lng_終了行
        ' セルの値を文字列として取得
        str_セル内の元のテキスト = CStr(Cells(lng_行番号, str_対象列).Value)
        
        ' 正規表現にマッチする箇所を全て取得
        Set obj_見つかった全ての候補 = obj_正規表現.Execute(str_セル内の元のテキスト)

        ' マッチする箇所がなければ次のセルへ
        If obj_見つかった全ての候補.Count = 0 Then
            GoTo 次のセルへ
        End If

        ' 新しい文字列を構築するための変数
        Dim lng_現在の処理位置 As Long
        Dim str_変更後の文字列 As String
        
        lng_現在の処理位置 = 1
        str_変更後の文字列 = ""

        ' 見つかった候補を一つずつ処理し、新しい文字列を組み立てる
        For Each obj_見つかった一つの候補 In obj_見つかった全ての候補
            ' マッチするまでの部分を「変更後の文字列」に追加
            str_変更後の文字列 = str_変更後の文字列 & Mid(str_セル内の元のテキスト, lng_現在の処理位置, obj_見つかった一つの候補.FirstIndex - lng_現在の処理位置 + 1)

            ' ハイフン前の数字を取得し、1加算
            lng_ハイフン前の数字 = CLng(obj_見つかった一つの候補.SubMatches(0))
            lng_ハイフン前の数字 = lng_ハイフン前の数字 + 1

            ' ハイフン後の数字を取得
            str_ハイフン後の数字 = obj_見つかった一つの候補.SubMatches(1)
            
            ' 新しい「置き換える文字列」を作成
            str_置き換える文字列 = CStr(lng_ハイフン前の数字) & "-" & str_ハイフン後の数字
            
            ' 「置き換える文字列」を「変更後の文字列」に追加
            str_変更後の文字列 = str_変更後の文字列 & str_置き換える文字列
            
            ' 次の処理開始位置を更新
            lng_現在の処理位置 = obj_見つかった一つの候補.FirstIndex + obj_見つかった一つの候補.Length + 1
        Next obj_見つかった一つの候補

        ' 最後のマッチ以降の部分を「変更後の文字列」に追加
        str_変更後の文字列 = str_変更後の文字列 & Mid(str_セル内の元のテキスト, lng_現在の処理位置)

        ' 変更後の文字列をセルに書き戻す
        Cells(lng_行番号, str_対象列).Value = str_変更後の文字列

次のセルへ:
    Next lng_行番号

    '--- 終了処理 ---
    Set obj_正規表現 = Nothing ' オブジェクトの解放
    MsgBox "置換処理が完了しました。", vbInformation
    
End Sub
