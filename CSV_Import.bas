Attribute VB_Name = "CSV_Import"
Option Explicit

' ============================================================
'  CSV Auto Import Macro
'  Target file: 0全科目月次ﾃﾞｰﾀ出力.xlsm
'  Function: Auto paste FreeWay CSV to target sheets
'
'  Sheet mapping:
'    (1) Current CSV (tax-excl) -> Sheets(4): 602全科目月次ﾃﾞｰﾀ出力（当期のみ）
'    (2) Tax CSV (tax-incl)     -> Sheets(6): 税込ﾃﾞｰﾀ専用
'    (3) 3year CSV              -> Sheets(1): 602全科目月次ﾃﾞｰﾀ出力（三期分）
' ============================================================

Sub CSV_Import()
    Dim p As String
    Dim ans As Integer
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Clear all 3 sheets first
    wb.Sheets(4).Cells.ClearContents
    wb.Sheets(6).Cells.ClearContents
    wb.Sheets(1).Cells.ClearContents

    ' (1) Current CSV (required)
    p = Application.GetOpenFilename("CSV,*.csv", , "(1) Select CSV")
    If p = "False" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    Call Paste(p, wb, 1)

    ' (2) Tax CSV (optional)
    ans = MsgBox("(2) Tax CSV?", vbYesNo)
    If ans = vbYes Then
        p = Application.GetOpenFilename("CSV,*.csv", , "(2) Select CSV")
        If p <> "False" Then Call Paste(p, wb, 2)
    End If

    ' (3) 3year CSV (optional)
    ans = MsgBox("(3) 3year CSV?", vbYesNo)
    If ans = vbYes Then
        p = Application.GetOpenFilename("CSV,*.csv", , "(3) Select CSV")
        If p <> "False" Then Call Paste(p, wb, 3)
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Done!"
End Sub

Private Sub Paste(p As String, wb As Workbook, mode As Integer)
    Dim ws As Worksheet
    Dim name As String
    If mode = 1 Then
        name = wb.Sheets(4).Name
    ElseIf mode = 2 Then
        name = wb.Sheets(6).Name
    ElseIf mode = 3 Then
        name = wb.Sheets(1).Name
    End If
    Set ws = wb.Worksheets(name)
    Dim csv As Workbook
    Set csv = Workbooks.Open(p)
    csv.Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
    csv.Close False
    wb.Activate
End Sub
