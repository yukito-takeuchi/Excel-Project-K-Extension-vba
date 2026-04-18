Option Explicit

Sub 手入力削除()
'
' 手入力削除 Macro
'

    ActiveWindow.SmallScroll Down:=12
    Range("G217").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G218").Select

    Columns("G:G").Select
    Range("G199").Activate
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    Selection.ClearContents
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 168
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 170
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 182
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 189
    ActiveWindow.ScrollRow = 190
    Range("G204").Select
End Sub




Sub 年間まとめて手入力削除()
'
' 年間まとめて手入力削除 Macro
'

'
    Sheets("①").Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    ActiveSheet.Next.Select
    Application.Run "'10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm'!手入力削除"
    Sheets(Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")).Select
    Sheets("⑫").Activate
    Application.Goto Reference:="R1C1"
    Sheets("BASE").Select
    Range("B28:O40").Select
    Selection.ClearContents
    Range("C28").Select
    Selection.FormulaR1C1 = "=R[12]C[-1]"
    Selection.Copy
    Range("D28:O28").Select
    Range("C28").Select
    Selection.Copy
    Range("D28:O28").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("B28").Select
    Application.CutCopyMode = False
End Sub

Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveWindow.SmallScroll Down:=12
    Range("G217").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G218").Select
End Sub
