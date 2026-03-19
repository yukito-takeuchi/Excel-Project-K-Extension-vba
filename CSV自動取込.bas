Attribute VB_Name = "CSV自動取込"
Option Explicit

' ============================================================
'  CSV自動取込マクロ
'  対象ファイル: 0全科目月次ﾃﾞｰﾀ出力.xlsm
'  機能: FreeWayから出力したCSVを対象シートへ自動貼り付け
' ============================================================

Sub CSV自動取込()

    Dim csvPath As String
    Dim thisWb As Workbook
    Dim ans As Integer

    Set thisWb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' -------------------------------------------------------
    ' ① 当期CSV（税抜）※必須
    ' -------------------------------------------------------
    csvPath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="① 当期CSVを選択してください（税抜）")

    ' キャンセルされた場合は処理中止
    If csvPath = "False" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "当期CSVが選択されませんでした。処理を中止します。", _
               vbExclamation, "中止"
        Exit Sub
    End If

    Call ImportCSV(csvPath, thisWb, "602全科目月次ﾃﾞｰﾀ出力（当期のみ）")

    ' -------------------------------------------------------
    ' ② 税込CSV（任意・スキップ可）
    ' -------------------------------------------------------
    ans = MsgBox("② 税込CSVを取り込みますか？" & Chr(10) & _
                 "スキップする場合は「いいえ」を押してください。", _
                 vbYesNo + vbQuestion, "税込CSV")

    If ans = vbYes Then
        csvPath = Application.GetOpenFilename( _
            FileFilter:="CSVファイル (*.csv),*.csv", _
            Title:="② 税込CSVを選択してください")

        If csvPath <> "False" Then
            Call ImportCSV(csvPath, thisWb, "税込ﾃﾞｰﾀ専用")
        End If
    End If

    ' -------------------------------------------------------
    ' ③ 三期分CSV（任意・スキップ可）
    ' -------------------------------------------------------
    ans = MsgBox("③ 三期分CSVを取り込みますか？" & Chr(10) & _
                 "スキップする場合は「いいえ」を押してください。", _
                 vbYesNo + vbQuestion, "三期分CSV")

    If ans = vbYes Then
        csvPath = Application.GetOpenFilename( _
            FileFilter:="CSVファイル (*.csv),*.csv", _
            Title:="③ 三期分CSVを選択してください")

        If csvPath <> "False" Then
            Call ImportCSV(csvPath, thisWb, "602全科目月次ﾃﾞｰﾀ出力（三期分）")
        End If
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "CSVの取り込みが完了しました。" & Chr(10) & Chr(10) & _
           "各ファイルを開いてグラフ生成ボタンを押してください。", _
           vbInformation, "取り込み完了"

End Sub

' ============================================================
'  CSV読み込み＆貼り付け（サブルーチン）
' ============================================================
Private Sub ImportCSV(csvPath As String, thisWb As Workbook, sheetName As String)

    Dim csvWb As Workbook
    Dim targetWs As Worksheet

    ' 対象シートを取得
    On Error Resume Next
    Set targetWs = thisWb.Worksheets(sheetName)
    On Error GoTo 0

    If targetWs Is Nothing Then
        MsgBox "シート「" & sheetName & "」が見つかりません。", _
               vbExclamation, "シートエラー"
        Exit Sub
    End If

    ' CSVを開く
    Set csvWb = Workbooks.Open(Filename:=csvPath)

    ' 全選択してコピー
    csvWb.Sheets(1).Cells.Select
    Selection.Copy

    ' 対象シートのA1に貼り付け
    targetWs.Range("A1").Select
    targetWs.Paste

    ' CSVを閉じる（保存しない）
    Application.CutCopyMode = False
    csvWb.Close SaveChanges:=False

    ' thisWbをアクティブに戻す
    thisWb.Activate

End Sub
