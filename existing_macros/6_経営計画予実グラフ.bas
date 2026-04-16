' ===== Module1（変数宣言のみ・処理なし） =====
    Public row As Integer         '表の左端の行数
    Public col As Integer         '表の左端の列数
    Public sh_name As String      '使用するシート名

    Public right As Integer
    Public left As Integer

    Dim touza_s As Long         '当座資産
    Dim ryudou_s As Long        '流動資産
    Dim kotei_s As Long         '固定資産
    Dim ryudou_f As Long        '流動負債
    Dim kotei_f As Long         '固定負債
    Dim shihon As Long          '資本

    Dim s_goukei As Long        '資産合計
    Dim f_goukei As Long        '負債合計
    Dim f_s_goukei As Long      '負債・資本合計

    Dim touza_hiritu As Double          '当座比率
    Dim ryudou_hiritu As Double         '流動比率
    Dim koteichoki_hiritu As Double     '固定長期適合率
    Dim kotei_hiritu As Double          '固定比率
    Dim jiko_hiritu As Double           '自己資本比率

    Dim cnt As Integer
    Dim err_msg As String

' ===== Module2（メインマクロ） =====


Option Explicit
Dim RowHeight As Single
Dim ob As Variant

Sub 経営計画図表作成()
   
  If Worksheets("入力ｼｰﾄ").Range("H21") >= 0 Then
    Worksheets("①PL").Activate
    Rows("4").RowHeight = Cells(4, 11)
    Rows("5").RowHeight = Cells(5, 11)
    Rows("6").RowHeight = Cells(6, 11)
  Else
    Worksheets("①PL").Activate
    Rows("16").RowHeight = Cells(16, 11)
    Rows("17").RowHeight = Cells(17, 11)
    Rows("18").RowHeight = Cells(18, 11)
  End If

   
  If Worksheets("入力ｼｰﾄ").Range("AC21") >= 0 Then
    Worksheets("②PL").Activate
    Rows("4").RowHeight = Cells(4, 11)
    Rows("5").RowHeight = Cells(5, 11)
    Rows("6").RowHeight = Cells(6, 11)
  Else
    Worksheets("②PL").Activate
    Rows("16").RowHeight = Cells(16, 11)
    Rows("17").RowHeight = Cells(17, 11)
    Rows("18").RowHeight = Cells(18, 11)
  End If
 
     
  If Worksheets("入力ｼｰﾄ").Range("AN21") >= 0 Then
    Worksheets("③PL").Activate
    Rows("4").RowHeight = Cells(4, 11)
    Rows("5").RowHeight = Cells(5, 11)
    Rows("6").RowHeight = Cells(6, 11)
  Else
    Worksheets("③PL").Activate
    Rows("16").RowHeight = Cells(16, 11)
    Rows("17").RowHeight = Cells(17, 11)
    Rows("18").RowHeight = Cells(18, 11)
  End If
 
  If Worksheets("入力ｼｰﾄ").Range("H33") > 0 Then
    Worksheets("経営計画 (繰欠あり)").Activate
  Else
    Worksheets("経営計画").Activate
  End If
 
  MsgBox "ラベル位置を調整してください。", vbInformation, "忘れずに"
 
 
End Sub