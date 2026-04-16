Dim I As Long
Dim k As Long
Dim h As Long
Dim word As String
Dim a As Long
Dim row_num As Long
Dim max_num As Long
Dim SearchWord As String
Dim check_name As String


Sub create_zukai()
   
   
    Dim cp_name As String
    Dim sh_name As String
    Dim FR As Range
   
    'シート名を取得
    sh_name = ActiveSheet.Name
   
    'コピー用シート名をセット
    If Left(sh_name, 2) = "累計" Then
        cp_name = "コピー用_累計"
    Else
        cp_name = "コピー用_単月"
    End If
   
    'シート名を取得
    sh_name = ActiveSheet.Name
   
    '前回のシートを削除してコピー
    msg = "現在の図解CFをコピーしますか？"
    Style = vbYesNo + vbQuestion
    Title = "要確認"
    msgRec = MsgBox(msg, Style, Title)
   
    'YESだった場合
    If msgRec = vbYes Then

        check_name = "前回分"
        '図解_累計シートをコピーする
        Sheet_Search2

    End If
   
    '図解がすでにあるかどうかチェック
    check_name = "図解_" & sh_name
    Sheet_Search1
   
    'コピー用シートをコピーする
    Worksheets(cp_name).Copy After:=ActiveSheet
    ActiveSheet.Name = "図解_" & sh_name
   
    '最大値を取得
    max_num = Cells(1, 61).Value
   
    '数字が入っていなかったら行を削除する
    '営業活動
    I = 18
    k = 3
    Do
        I = I + 1
        word = Left(Cells(I, k).Value, 2)
        '削除する
        delete_row
             
    Loop Until word = "小計"
   
    '小計と営業活動により調達した純キャッシュの括弧をはずして結合する
    If Cells(I, 12).Value <> "" Then
        Range(Cells(I, 56), Cells(I, 60)).ClearContents
'        Cells(I, 60).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  ▲" & Format(Abs(Cells(I, 12)), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(I, 11), Cells(I, 15)).ClearContents
        Range(Cells(I, 11), Cells(I, 15)).UnMerge

        Cells(I, 11) = cash
        Range(Cells(I, 11), Cells(I, 35)).Merge
        Cells(I, 11).HorizontalAlignment = xlLeft
        Cells(I, 11).VerticalAlignment = xlCenter
       
    Else
        Range(Cells(I, 11), Cells(I, 15)).ClearContents
'        Cells(I, 15).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  +" & Format(Cells(I, 57), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(I, 56), Cells(I, 60)).ClearContents
        Range(Cells(I, 56), Cells(I, 60)).UnMerge
       
        Cells(I, 60) = cash
        Range(Cells(I, 36), Cells(I, 60)).Merge
        Cells(I, 36).HorizontalAlignment = xlRight
        Cells(I, 36).VerticalAlignment = xlCenter
       
    End If

   
    If Cells(I + 3, 12).Value <> "" Then
        Range(Cells(I + 3, 56), Cells(I + 3, 60)).ClearContents
'        Cells(I + 3, 60).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  ▲" & Format(Abs(Cells(I + 3, 12)), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(I + 3, 11), Cells(I + 3, 15)).ClearContents
        Range(Cells(I + 3, 11), Cells(I + 3, 15)).UnMerge

        Cells(I + 3, 11) = cash
        Range(Cells(I + 3, 11), Cells(I + 3, 35)).Merge
        Cells(I + 3, 11).HorizontalAlignment = xlLeft
        Cells(I + 3, 11).VerticalAlignment = xlCenter
       
    Else
        Range(Cells(I + 3, 11), Cells(I + 3, 15)).ClearContents
'        Cells(I + 3, 15).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  +" & Format(Cells(I + 3, 57), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(I + 3, 56), Cells(I + 3, 60)).ClearContents
        Range(Cells(I + 3, 56), Cells(I + 3, 60)).UnMerge
       
        Cells(I + 3, 60) = cash
        Range(Cells(I + 3, 36), Cells(I + 3, 60)).Merge
        Cells(I + 3, 36).HorizontalAlignment = xlRight
        Cells(I + 3, 36).VerticalAlignment = xlCenter
       
       
    End If
   
   
    '投資活動
    k = 2
    I = I + 11
    Do
        I = I + 1
        word = Left(Cells(I, k).Value, 3)
        '削除する
        delete_row
       
    Loop Until word = "その他"
   
    '投資活動に使用した純キャッシュの括弧をはずす
    SearchWord = "投資活動に使用した"
    With ActiveSheet.Cells
    Set FR = .Find(SearchWord)
   
    If Cells(FR.Row, 12).Value <> "" Then
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).ClearContents
'        Cells(FR.Row, 60).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  ▲" & Format(Abs(Cells(FR.Row, 12)), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).ClearContents
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).UnMerge

        Cells(FR.Row, 11) = cash
        Range(Cells(FR.Row, 11), Cells(FR.Row, 35)).Merge
        Cells(FR.Row, 11).HorizontalAlignment = xlLeft
        Cells(FR.Row, 11).VerticalAlignment = xlCenter
       
    Else
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).ClearContents
'        Cells(FR.Row, 15).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  +" & Format(Cells(FR.Row, 57), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).ClearContents
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).UnMerge
       
        Cells(FR.Row, 60) = cash
        Range(Cells(FR.Row, 36), Cells(FR.Row, 60)).Merge
        Cells(FR.Row, 36).HorizontalAlignment = xlRight
        Cells(FR.Row, 36).VerticalAlignment = xlCenter
       
    End If
   
    End With
    Set FR = Nothing
   
    '財務活動
    k = 2
   
    I = I + 18
    Do
        I = I + 1
        word = Left(Cells(I, k).Value, 2)
        delete_row
    Loop Until word = "貸付"
   
   
    '財務活動に使用した純キャッシュの括弧をはずす
    SearchWord = "財務活動に使用した"
    With ActiveSheet.Cells
    Set FR = .Find(SearchWord)
   
    If Cells(FR.Row, 12).Value <> "" Then
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).ClearContents
'        Cells(FR.Row, 60).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  ▲" & Format(Abs(Cells(FR.Row, 12)), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).ClearContents
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).UnMerge

        Cells(FR.Row, 11) = cash
        Range(Cells(FR.Row, 11), Cells(FR.Row, 35)).Merge
        Cells(FR.Row, 11).HorizontalAlignment = xlLeft
        Cells(FR.Row, 11).VerticalAlignment = xlCenter
       
    Else
        Range(Cells(FR.Row, 11), Cells(FR.Row, 15)).ClearContents
'        Cells(FR.Row, 15).ClearContents
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  +" & Format(Cells(FR.Row, 57), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).ClearContents
        Range(Cells(FR.Row, 56), Cells(FR.Row, 60)).UnMerge
       
        Cells(FR.Row, 60) = cash
        Range(Cells(FR.Row, 36), Cells(FR.Row, 60)).Merge
        Cells(FR.Row, 36).HorizontalAlignment = xlRight
        Cells(FR.Row, 36).VerticalAlignment = xlCenter
       
    End If
   
    End With
    Set FR = Nothing
   
   

   
    '色付けをする
    For a = 1 To 30
        SearchWord = Worksheets("work").Cells(a, 1).Value
        With ActiveSheet.Cells
        Set FR = .Find(SearchWord)
        If Not FR Is Nothing Then
            row_num = FR.Row
            color
        End If
        End With
        Set FR = Nothing
    Next
   
    '最大値を削除
    Cells(1, 61).ClearContents
   
Cells(1, 1).Select

End Sub


'数字が入っていなかった場合に削除する
Sub delete_row()
   
    '項目があったらチェックする
    If Cells(I, k).Value <> "" Then
        '数字が入っていなかったら対象行とその下の行を削除する
        If Cells(I, 12).Value = "" And Cells(I, 57).Value = "" Then
            Rows(I).Delete
            Rows(I).Delete
            '削除したあとに行数を-1する
            I = I - 1
        End If
    End If
   
   
End Sub


'数字の割合に応じて色付けをする
Sub color()
   
    Dim kingaku As Long
    Dim cnt As Double
    Dim z As Long
    Dim color As Long
    Dim cash As String
   
    'プラスかマイナスか判定
    If Cells(row_num, 12).Value = "" Then
    'プラスだった場合
        'マイナスの括弧をはずす
        Range(Cells(row_num, 11), Cells(row_num, 15)).ClearContents
'        Cells(row_num, 15).ClearContents
       
        '色をセット
        If SearchWord = "フリー純" Then
            color = 16777093
        Else
            color = 12611584
        End If
       
       
        '金額を取得
        kingaku = Abs(Cells(row_num, 57).Value)
       
        '色づけ個数を計算
        cnt = Round(60 * (kingaku / max_num), 0)
        '0だった場合１に設定
        If cnt = 0 Then
            cnt = 1
        End If
       
        '1～20個だった場合
        If cnt <= 20 And SearchWord <> "現預金「期首」残高" And SearchWord <> "現預金「期末」残高" And SearchWord <> "現預金「月初」残高" And SearchWord <> "現預金「月末」残高" Then
            z = cnt
            Range(Cells(row_num, 36), Cells(row_num, 35 + cnt)).Select
            Selection.Interior.color = color

        '21～40個だった場合
        ElseIf 20 < cnt And cnt <= 40 And SearchWord <> "現預金「期首」残高" And SearchWord <> "現預金「期末」残高" And SearchWord <> "現預金「月初」残高" And SearchWord <> "現預金「月末」残高" Then
           
            '１行挿入する
            Rows(row_num + 1).EntireRow.Insert
           
            '個数を半分にセットする
            z = Round(cnt / 2, 0)
            Range(Cells(row_num, 36), Cells(row_num + 1, 35 + z)).Select
            Selection.Interior.color = color
           
        '41個以上だった場合
        ElseIf cnt > 40 Or SearchWord = "現預金「期首」残高" Or SearchWord = "現預金「期末」残高" And SearchWord <> "現預金「月初」残高" And SearchWord <> "現預金「月末」残高" Then
       
            '２行挿入する
            Rows(row_num + 1).EntireRow.Insert
            Rows(row_num + 1).EntireRow.Insert
           
            '個数を3/1にセットする
            z = Round(cnt / 3, 0)
           
            '現金残が0ではなくてZが0の場合は1にする
            If Cells(row_num, 57) <> 0 And z = 0 Then
                z = 1
            End If
           
            If z <> 0 Then
           
                Range(Cells(row_num, 36), Cells(row_num + 2, 35 + z)).Select
                Selection.Interior.color = color
           
            End If
       
       
        End If
       
        '数字を位置を調整する
        '数字をセットする
        If Cells(row_num, 57) <> 0 Then
            cash = "(  +" & Format(Cells(row_num, 57), "#,#") & "  )"
        Else
            cash = "'(　0　)"
        End If

        '数字を削除して結合を外す
        Range(Cells(row_num, 56), Cells(row_num, 60)).ClearContents
        Range(Cells(row_num, 56), Cells(row_num, 60)).UnMerge
       
        Cells(row_num, 36 + z) = cash
        Range(Cells(row_num, 36 + z), Cells(row_num, 60)).Merge
        Cells(row_num, 36 + z).HorizontalAlignment = xlLeft
        Cells(row_num, 36 + z).VerticalAlignment = xlCenter
       
    Else
    'マイナスだった場合
       
        'プラスの括弧をはずす
        Range(Cells(row_num, 56), Cells(row_num, 60)).ClearContents
'        Cells(row_num, 60).ClearContents
       
        '色をセット
        If SearchWord = "フリー純" Then
            color = 16777093
        Else
            color = 255
        End If
       
        '金額を取得
        kingaku = Abs(Cells(row_num, 12).Value)
       
        '色づけ個数を計算
        cnt = Round(60 * (kingaku / max_num), 0)
        '0だった場合1に設定
        If cnt = 0 Then
            cnt = 1
        End If
       
        '1～20個だった場合
        If cnt <= 20 Then
            z = cnt
            Range(Cells(row_num, 35), Cells(row_num, 36 - cnt)).Select
            Selection.Interior.color = color

        '21～40個だった場合
        ElseIf 20 < cnt And cnt <= 40 Then
           
            '１行挿入する
            Rows(row_num + 1).EntireRow.Insert
           
            '個数を半分にセットする
            z = Round(cnt / 2, 0)
            Range(Cells(row_num, 35), Cells(row_num + 1, 36 - z)).Select
            Selection.Interior.color = color
           
        '41個以上だった場合
        ElseIf cnt > 40 Then
       
            '２行挿入する
            Rows(row_num + 1).EntireRow.Insert
            Rows(row_num + 1).EntireRow.Insert
           
            '個数を3/1にセットする
            z = Round(cnt / 3, 0)
            Range(Cells(row_num, 35), Cells(row_num + 2, 36 - z)).Select
            Selection.Interior.color = color
       
        End If
       
       
        '数字を位置を調整する
        '数字をセットする
        cash = "(  ▲" & Format(Abs(Cells(row_num, 12)), "#,#") & "  )"
        '数字を削除して結合を外す
        Range(Cells(row_num, 11), Cells(row_num, 15)).ClearContents
        Range(Cells(row_num, 11), Cells(row_num, 15)).UnMerge

        Cells(row_num, 11) = cash
        Range(Cells(row_num, 11), Cells(row_num, 35 - z)).Merge
        Cells(row_num, 11).HorizontalAlignment = xlRight
        Cells(row_num, 11).VerticalAlignment = xlCenter
    End If

End Sub

'図解_累計シートがあるかどうかチェック
Sub Sheet_Search1()

    Dim N As Integer, Check As Integer, I As Integer
   
On Error Resume Next
    Check = 1
    N = Worksheets.Count
    For I = 1 To N
        If Worksheets(I).Name = (check_name) Then
            MsgBox check_name & "シートを削除してから実行してください", vbCritical
            End
        End If
    Next
End Sub

'前回シートをコピー
Sub Sheet_Search2()

    Dim N As Integer, Check As Integer, I As Integer
   
On Error Resume Next
    Check = 1
    N = Worksheets.Count
    For I = 1 To N
        '検索するシートが見つかった場合そのシートを削除する
        If Worksheets(I).Name = (check_name) Then
            Application.DisplayAlerts = False
            Worksheets(check_name).Delete
            Application.DisplayAlerts = True
        End If
    Next

    '図解_累計シートのシート名を前回分に変更する
    Worksheets("図解_累計").Name = "前回分"


'    '検索するシートが前回分だった場合、図解_累計シートをコピーして図解_累計シートを削除する
'    If sh_name = "前回分" Then
'        Worksheets(sh_name).Cells.Copy
'        Worksheets(sh_name).Cells.PasteSpecial Paste:=xlPasteValues
'        Worksheets(sh_name).Copy After:=Worksheets(sh_name)
'        ActiveSheet.Cells(1, 1).Select
'        ActiveSheet.Name = "前回分"
'
'        Application.DisplayAlerts = False
'        Worksheets("図解_累計").Delete
'        Application.DisplayAlerts = True
'
'    End If
   
End Sub