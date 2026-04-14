Attribute VB_Name = "CSV自動取込"
Option Explicit

' ============================================================
' 定数定義：シートインデックス（マジックナンバー排除）
' ============================================================
Private Const IDX_三期分  As Integer = 1   ' 三期分データ貼り付け先
Private Const IDX_当期    As Integer = 4   ' 当期CSVの貼り付け先（602当期のみ）
Private Const IDX_税込    As Integer = 6   ' 税込データ専用シート

' ============================================================
' 定数定義：消費税方式
' ============================================================
Private Const MODE_税抜 As String = "税抜"
Private Const MODE_税込 As String = "税込"

' ============================================================
' メインマクロ：CSV自動取込
' ============================================================
Sub CSV_Import()

    Dim wb          As Workbook
    Dim taxMode     As String
    Dim csvPath     As String
    Dim ans         As Integer
    Dim doTaxIn     As Boolean      ' 税込CSVを取り込むか
    Dim doThreeYear As Boolean      ' 三期分CSVを取り込むか
    Dim rows1       As Long         ' 当期CSV 取り込み行数
    Dim rows2       As Long         ' 税込CSV 取り込み行数
    Dim rows3       As Long         ' 三期分CSV 取り込み行数
    Dim msg         As String

    Set wb = ThisWorkbook

    ' エラーハンドラを設定（予期せぬエラーで状態が壊れないよう保護）
    On Error GoTo ErrHandler

    ' ----------------------------------------------------------
    ' STEP 0：消費税方式を選択
    ' ----------------------------------------------------------
    taxMode = 消費税方式を選択()
    If taxMode = "" Then Exit Sub   ' キャンセル時は何もしない

    ' ----------------------------------------------------------
    ' STEP 1：取り込むCSVをユーザーに確認（クリア前に判断する）
    ' ※ スキップしたシートのデータは消えないよう、先に意向を確認する
    ' ----------------------------------------------------------

    ' 税込CSV（税抜モードのみ対象。税込モードでは数式連動のため不要）
    If taxMode = MODE_税抜 Then
        ans = MsgBox("② 税込CSVを取り込みますか？" & Chr(13) & _
                     "（「いいえ」でスキップ。既存データは保持されます）", _
                     vbYesNo + vbQuestion, "② 税込CSV の取り込み確認")
        doTaxIn = (ans = vbYes)
    Else
        ' 税込モード：Sheet(6)への貼り付けは行わない（数式で自動連動）
        doTaxIn = False
    End If

    ' 三期分CSV
    ans = MsgBox("③ 三期分CSVを取り込みますか？" & Chr(13) & _
                 "（「いいえ」でスキップ。既存データは保持されます）", _
                 vbYesNo + vbQuestion, "③ 三期分CSV の取り込み確認")
    doThreeYear = (ans = vbYes)

    ' ----------------------------------------------------------
    ' STEP 2：実行前確認ダイアログ
    ' ----------------------------------------------------------
    msg = "以下のシートの既存データをクリアして取り込みを開始します。よろしいですか？" & Chr(13) & Chr(13)
    msg = msg & "・消費税方式：" & taxMode & Chr(13)
    msg = msg & "・① 当期CSV → シート(" & IDX_当期 & ")" & Chr(13)
    If doTaxIn Then      msg = msg & "・② 税込CSV → シート(" & IDX_税込 & ")" & Chr(13)
    If doThreeYear Then  msg = msg & "・③ 三期分CSV → シート(" & IDX_三期分 & ")" & Chr(13)

    ans = MsgBox(msg, vbYesNo + vbQuestion, "実行確認")
    If ans = vbNo Then Exit Sub

    ' ----------------------------------------------------------
    ' STEP 3：画面更新を停止（処理高速化・ちらつき防止）
    ' ----------------------------------------------------------
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ----------------------------------------------------------
    ' STEP 4：シート存在チェック → クリア（取り込むシートのみ）
    ' ----------------------------------------------------------
    Application.StatusBar = "既存データをクリア中..."

    ' 当期シート（必須）
    If Not シート存在確認(wb, IDX_当期) Then GoTo Cleanup
    wb.Sheets(IDX_当期).Cells.ClearContents

    ' 税込シート（任意）
    If doTaxIn Then
        If Not シート存在確認(wb, IDX_税込) Then GoTo Cleanup
        wb.Sheets(IDX_税込).Cells.ClearContents
    End If

    ' 三期分シート（任意）
    If doThreeYear Then
        If Not シート存在確認(wb, IDX_三期分) Then GoTo Cleanup
        wb.Sheets(IDX_三期分).Cells.ClearContents
    End If

    ' ----------------------------------------------------------
    ' STEP 5：① 当期CSV（必須）
    ' ----------------------------------------------------------
    Application.StatusBar = "① 当期CSV を選択してください..."
    csvPath = Application.GetOpenFilename("CSV,*.csv", , "① 当期CSV を選択")

    If csvPath = "False" Then
        ' キャンセル → スキップして続行するか確認
        Application.ScreenUpdating = True   ' ダイアログ表示のため一時的に再開
        ans = MsgBox("① 当期CSVの選択がキャンセルされました。" & Chr(13) & _
                     "スキップして残りの処理を続けますか？（「いいえ」で中止）", _
                     vbYesNo + vbExclamation, "確認")
        Application.ScreenUpdating = False
        If ans = vbNo Then GoTo Cleanup
    Else
        Application.StatusBar = "① 当期CSV を貼り付け中..."
        rows1 = CSVを貼付(csvPath, wb, IDX_当期)
        If rows1 < 0 Then
            MsgBox "① 当期CSVの取り込みに失敗しました。処理を中止します。", vbCritical
            GoTo Cleanup
        End If
    End If

    ' ----------------------------------------------------------
    ' STEP 6：② 税込CSV（税抜モード・任意）
    ' ----------------------------------------------------------
    If doTaxIn Then
        Application.StatusBar = "② 税込CSV を選択してください..."
        csvPath = Application.GetOpenFilename("CSV,*.csv", , "② 税込CSV を選択")

        If csvPath = "False" Then
            MsgBox "② 税込CSVの選択がキャンセルされました。スキップして続行します。", vbInformation
        Else
            Application.StatusBar = "② 税込CSV を貼り付け中..."
            rows2 = CSVを貼付(csvPath, wb, IDX_税込)
            If rows2 < 0 Then
                MsgBox "② 税込CSVの取り込みに失敗しました。スキップして続行します。", vbExclamation
                rows2 = 0
            End If
        End If
    End If

    ' ----------------------------------------------------------
    ' STEP 7：③ 三期分CSV（任意）
    ' ----------------------------------------------------------
    If doThreeYear Then
        Application.StatusBar = "③ 三期分CSV を選択してください..."
        csvPath = Application.GetOpenFilename("CSV,*.csv", , "③ 三期分CSV を選択")

        If csvPath = "False" Then
            MsgBox "③ 三期分CSVの選択がキャンセルされました。スキップして続行します。", vbInformation
        Else
            Application.StatusBar = "③ 三期分CSV を貼り付け中..."
            rows3 = CSVを貼付(csvPath, wb, IDX_三期分)
            If rows3 < 0 Then
                MsgBox "③ 三期分CSVの取り込みに失敗しました。スキップして続行します。", vbExclamation
                rows3 = 0
            End If
        End If
    End If

    ' ----------------------------------------------------------
    ' STEP 8：完了メッセージ（取り込み行数を表示）
    ' ----------------------------------------------------------
    msg = "取り込みが完了しました！" & Chr(13) & Chr(13)
    msg = msg & "消費税方式：" & taxMode & Chr(13)
    If rows1 > 0 Then msg = msg & "① 当期CSV：" & rows1 & " 行 取り込み完了" & Chr(13)
    If rows2 > 0 Then msg = msg & "② 税込CSV：" & rows2 & " 行 取り込み完了" & Chr(13)
    If rows3 > 0 Then msg = msg & "③ 三期分CSV：" & rows3 & " 行 取り込み完了" & Chr(13)

    MsgBox msg, vbInformation, "完了"
    GoTo Cleanup

    ' ----------------------------------------------------------
    ' エラーハンドラ：予期せぬエラーをキャッチし、状態を必ず復元する
    ' ----------------------------------------------------------
ErrHandler:
    MsgBox "予期せぬエラーが発生しました。処理を中止します。" & Chr(13) & Chr(13) & _
           "エラー番号：" & Err.Number & Chr(13) & _
           "内容：" & Err.Description, vbCritical, "エラー"

Cleanup:
    ' Application の状態を必ず復元する（エラー時も保証）
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    wb.Activate

End Sub

' ============================================================
' 消費税方式を選択するダイアログ
' 戻り値："税抜" / "税込" / ""（キャンセル）
' ============================================================
Private Function 消費税方式を選択() As String

    Dim ans As Integer

    ans = MsgBox("消費税方式を選択してください。" & Chr(13) & Chr(13) & _
                 "【はい】→ 税抜方式（当期CSV + 税込CSV を別々に取り込む）" & Chr(13) & _
                 "【いいえ】→ 税込方式（当期CSVのみ。税込列は数式で自動連動）", _
                 vbYesNoCancel + vbQuestion, "消費税方式の選択")

    Select Case ans
        Case vbYes:     消費税方式を選択 = MODE_税抜
        Case vbNo:      消費税方式を選択 = MODE_税込
        Case vbCancel:  消費税方式を選択 = ""   ' キャンセル
    End Select

End Function

' ============================================================
' CSVをシートに貼り付けるFunction
' 引数：csvPath  = CSVファイルのフルパス
'       wb       = 対象ブック
'       sheetIdx = 貼り付け先シートのインデックス番号
' 戻り値：取り込んだ行数（空ファイル=0、エラー=-1）
' ============================================================
Private Function CSVを貼付(csvPath As String, wb As Workbook, sheetIdx As Integer) As Long

    Dim ws       As Worksheet
    Dim csvWb    As Workbook
    Dim rowCount As Long

    On Error GoTo PasteError

    Set ws = wb.Sheets(sheetIdx)

    ' CSVファイルを開く
    Set csvWb = Workbooks.Open(csvPath)

    ' 空ファイルチェック（UsedRangeが1セルかつ空の場合）
    With csvWb.Sheets(1).UsedRange
        rowCount = .Rows.Count
        If rowCount <= 1 And Trim(csvWb.Sheets(1).Cells(1, 1).Value) = "" Then
            csvWb.Close False
            MsgBox "選択されたCSVファイルにデータがありません。スキップします。" & Chr(13) & _
                   "ファイル：" & csvPath, vbExclamation
            CSVを貼付 = 0
            Exit Function
        End If
    End With

    ' 貼り付け実行
    csvWb.Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
    csvWb.Close False
    wb.Activate

    CSVを貼付 = rowCount
    Exit Function

PasteError:
    ' エラー発生時はCSVを閉じてから -1 を返す
    On Error Resume Next
    If Not csvWb Is Nothing Then csvWb.Close False
    On Error GoTo 0

    ' エラー種別に応じてメッセージを分岐
    ' 1004 は Excel 汎用エラー。文字コード問題（Shift-JIS 以外）でも発生しやすいため案内を追加
    If Err.Number = 1004 Then
        MsgBox "CSVファイルの読み込みに失敗しました（エラー 1004）。" & Chr(13) & Chr(13) & _
               "選択されたCSVファイルの文字コードがShift-JIS以外の可能性があります。" & Chr(13) & Chr(13) & _
               "対処方法：" & Chr(13) & _
               "1. Excelでそのファイルを開く（文字コードをUTF-8で指定）" & Chr(13) & _
               "2. 別名保存でCSVを保存し直す" & Chr(13) & _
               "3. 再度このマクロを実行してください", _
               vbExclamation, "読み込みエラー（文字コードの可能性あり）"
    Else
        MsgBox "CSVファイルの読み込み中にエラーが発生しました。" & Chr(13) & Chr(13) & _
               "エラー番号：" & Err.Number & Chr(13) & _
               "内容：" & Err.Description, _
               vbCritical, "読み込みエラー"
    End If

    CSVを貼付 = -1

End Function

' ============================================================
' シートがインデックス番号で存在するか確認する関数
' 存在しない場合はエラーメッセージを表示してFalseを返す
' ============================================================
Private Function シート存在確認(wb As Workbook, idx As Integer) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Sheets(idx)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Application状態を一時的に復元してダイアログを表示
        Application.ScreenUpdating = True
        MsgBox "シート(" & idx & ")が見つかりません。" & Chr(13) & Chr(13) & _
               "シート名が変更されていないか確認してください。" & Chr(13) & _
               "処理を中止します。", vbCritical, "シートエラー"
        Application.ScreenUpdating = False
        シート存在確認 = False
    Else
        シート存在確認 = True
    End If

End Function
