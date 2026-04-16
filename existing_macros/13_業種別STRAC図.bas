Option Explicit
Dim RowHeight As Single

Sub ストラック図作成()
   
  Worksheets("黒字PL").Activate
  Rows("4").RowHeight = cells(4, 14)
  Rows("5").RowHeight = cells(5, 14)
  Rows("6").RowHeight = cells(6, 14)
   
  Worksheets("黒字BS").Activate
  Rows("4").RowHeight = cells(4, 9)
  Rows("5").RowHeight = cells(5, 9)
  Rows("11").RowHeight = cells(11, 9)
  Rows("12").RowHeight = cells(12, 9)
  Rows("13").RowHeight = cells(13, 9)
 
  If Worksheets("入力").Range("c3") > 0 Then
   
    Worksheets("優良PL").Activate
    Rows("4").RowHeight = cells(4, 14)
    Rows("5").RowHeight = cells(5, 14)
    Rows("6").RowHeight = cells(6, 14)
   
    Worksheets("優良BS").Activate
    Rows("4").RowHeight = cells(4, 9)
    Rows("5").RowHeight = cells(5, 9)
    Rows("11").RowHeight = cells(11, 9)
    Rows("12").RowHeight = cells(12, 9)
    Rows("13").RowHeight = cells(13, 9)
   
    Sheets("業種STRAC図").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$R$70,$S$1:$AJ$70"
    Range("A51").Select
   
  Else
   
    Sheets("業種STRAC図").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$R$70"
    Range("A51").Select
 
  End If

  MsgBox "ラベル位置を調整してください。", vbInformation, "忘れずに"
 
End Sub