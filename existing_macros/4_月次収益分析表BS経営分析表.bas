Sub ストラック図作成()
   
     
  If Worksheets("STRAC入力").Range("B23") >= 0 Then
    Worksheets("単PL").Activate
    Rows("4").RowHeight = Cells(4, 75)
    Rows("5").RowHeight = Cells(5, 75)
    Rows("6").RowHeight = Cells(6, 75)
    Rows("7").RowHeight = Cells(7, 75)
  Else
    Worksheets("単PL").Activate
     With ActiveSheet
     For Each ob In .DrawingObjects
     If Not Intersect(ob.TopLeftCell, .Range("AO16:BT18")) Is Nothing Then
     ob.Delete
     End If
     Next
     End With
    Rows("16").RowHeight = Cells(16, 75)
    Rows("17").RowHeight = Cells(17, 75)
    Rows("18").RowHeight = Cells(18, 75)
    Rows("22").RowHeight = Cells(22, 75)
    Rows("23").RowHeight = Cells(23, 75)
     Range("AO22:BT23").Select
     Selection.Copy
     Range("AO17:BT17").Select
     ActiveSheet.Pictures.Paste(Link:=True).Select
  End If
 
  If Worksheets("STRAC入力").Range("c23") >= 0 Then
    Worksheets("累PL").Activate
    Rows("4").RowHeight = Cells(4, 75)
    Rows("5").RowHeight = Cells(5, 75)
    Rows("6").RowHeight = Cells(6, 75)
    Rows("7").RowHeight = Cells(7, 75)
  Else
    Worksheets("累PL").Activate
    With ActiveSheet
    For Each ob In .DrawingObjects
    If Not Intersect(ob.TopLeftCell, .Range("AO16:BT18")) Is Nothing Then
    ob.Delete
    End If
    Next
    End With
    Rows("16").RowHeight = Cells(16, 75)
    Rows("17").RowHeight = Cells(17, 75)
    Rows("18").RowHeight = Cells(18, 75)
    Rows("22").RowHeight = Cells(22, 75)
    Rows("23").RowHeight = Cells(23, 75)
     Range("AO22:BT23").Select
     Selection.Copy
     Range("AO17:BT17").Select
     ActiveSheet.Pictures.Paste(Link:=True).Select
  End If
 
  If Worksheets("STRAC入力").Range("b34") >= 0 Then
    Worksheets("BS").Activate
    Rows("4").RowHeight = Cells(4, 17)
    Rows("5").RowHeight = Cells(5, 17)
    Rows("6").RowHeight = Cells(6, 17)
    Rows("12").RowHeight = Cells(12, 17)
    Rows("13").RowHeight = Cells(13, 17)
    Rows("14").RowHeight = Cells(14, 17)
    Rows("15").RowHeight = Cells(15, 17)
  Else
    Worksheets("BS").Activate
    Rows("25").RowHeight = Cells(25, 17)
    Rows("26").RowHeight = Cells(26, 17)
    Rows("27").RowHeight = Cells(27, 17)
    Rows("28").RowHeight = Cells(28, 17)
    Rows("35").RowHeight = Cells(35, 17)
     Rows("36").RowHeight = Cells(36, 17)
    Rows("37").RowHeight = Cells(37, 17)
  End If
 
    Sheets("STRAC図").Select
    Range("a1").Select
 
    MsgBox "ラベル位置を調整してください。", vbInformation, "忘れずに"
 
End Sub