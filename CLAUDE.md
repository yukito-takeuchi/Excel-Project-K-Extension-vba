# 黒永様 財務レポート自動化システム

## 参照ドキュメント（毎セッション必読）
- **要件定義書** → `docs/要件整理.md`

> 実装判断に迷ったら必ず上記ファイルを参照すること。
> 要件定義書に記載のない機能は独自判断で追加しない。

## プロジェクト概要

FreeWay（会計ソフト）のCSV → `0全科目月次データ出力.xlsm` の各シートへ自動貼り付けするVBAマクロを実装する。

**最重要制約：既存ファイル（1〜12）は一切変更しない。編集対象は `CSV自動取込.bas` のみ。**

## 既存システムの仕組み（必読）

```
FreeWay → CSV出力
    ↓ マクロで自動貼り付け（今回の実装範囲）
0全科目月次データ出力.xlsm
├ Sheet(1)「602三期分」     ← 三期分CSVの貼り付け先
├ Sheet(4)「602当期のみ」   ← 当期CSVの貼り付け先（必須）
└ Sheet(6)「税込データ専用」← 税込CSVの貼り付け先（税抜方式のみ）
    ↓ Excelの数式リンク（マクロ不要・自動）
1表紙.xlsx ～ 12決算書図表.xlsm（全ファイルが自動更新）
```

| 方式 | 三期分CSV | 当期CSV | 税込CSV |
|------|----------|---------|---------|
| 税抜方式 | Sheet(1)へ | Sheet(4)へ | Sheet(6)へ |
| 税込方式 | Sheet(1)へ | Sheet(4)へ | **スキップ**（数式で自動連動） |

## ディレクトリ構成

```
/
├ CLAUDE.md                  ← 本ファイル
├ docs/要件整理.md            ← 要件定義書
├ CLAUDE.md                  ← 本ファイル（vba/ 直下に配置・git管理）
├ docs/
│   └ 要件整理.md            ← 要件定義書
├ CSV_Import.bas             ← 英語版（参考のみ・編集不要）
├ CSV自動取込.bas            ← 実装対象（このファイルのみ編集する）
../demo/
│   └ index.html             ← クライアント向けデモUI（demo/.git で管理）
../*.xlsm / *.xlsx           ← 実際のExcelファイル群（変更禁止）
```

## 開発ルール（必読）

- シートインデックスは必ずConst定数を使う（直書き禁止）
- コメントはすべて日本語で記述する
- エラー時は必ず `Cleanup:` に進んで `ScreenUpdating / DisplayAlerts / StatusBar` を復元する
- `CSV自動取込.bas` は **全面書き直し**。現在の旧版（59行）は参考程度に見るだけでよい
- タスク完了後は本ファイルの該当チェックボックスを `[x]` に更新する
- セッション終了時は「引き継ぎメモ」に作業内容と次にやることを追記する

---

## 進捗状況

### ✅ 完了済み
（完了したタスクをここに移動する）

### 🔴 未着手（実装必須）

---

## 【実装セッションA】VBAマクロ全面実装

**指示：以下のTask A-1〜A-7を番号順に実装してください。全部で1つの `.bas` ファイルを完成させます。完了後に各チェックボックスを `[x]` にして、引き継ぎメモを更新してください。**

> 実装対象：`vba/CSV自動取込.bas`（旧版を全面書き直し）
> テスト環境：Macで実装 → Windowsで動作確認

---

### 🔵 Task A-1：ファイル冒頭・定数定義（最初に実装）

**背景：** 現在 `Sheets(4)` のようなマジックナンバーが直書きされており、シート順が変わると全壊する。Const定数で一元管理する。

- [ ] **A-1：モジュール宣言とConst定数をファイル冒頭に記述する**
  - 対象：`vba/CSV自動取込.bas` の先頭
  - 実装内容：
    ```vb
    Attribute VB_Name = "CSV自動取込"
    Option Explicit

    ' ============================================================
    '  CSV自動取込マクロ
    '  対象ファイル：0全科目月次ﾃﾞｰﾀ出力.xlsm
    '  FreeWayから出力したCSVを対象シートへ自動貼り付けする
    ' ============================================================

    ' シートインデックス定数（シート順が変わった場合ここだけ変更する）
    Private Const IDX_三期分 As Integer = 1
    Private Const IDX_当期   As Integer = 4
    Private Const IDX_税込   As Integer = 6
    ```
  - 完了条件：`Option Explicit` とConst定数3つが定義されている

---

### 🔵 Task A-2：メインSubの骨格・消費税方式選択

**背景：** 税抜/税込の方式により、取り込むCSVの数と貼り付け先が変わる。実行時にMsgBoxで選択させる。

- [ ] **A-2：`CSV取込` Sub の骨格と消費税方式選択ロジックを実装する**
  - 対象：`vba/CSV自動取込.bas`
  - 実装内容：
    ```vb
    Sub CSV取込()
        Dim wb       As Workbook  : Set wb = ThisWorkbook
        Dim taxMode  As String    ' "nuki"=税抜, "komi"=税込
        Dim skip三期  As Boolean
        Dim skip税込  As Boolean
        Dim rows当期  As Long
        Dim rows税込  As Long
        Dim rows三期  As Long

        ' --- 消費税方式の選択 ---
        ' Yes=税抜 / No=税込 / Cancel=中止
        Dim taxAns As Integer
        taxAns = MsgBox("消費税方式を選択してください。" & vbCrLf & _
                        "【はい】税抜方式　　【いいえ】税込方式", _
                        vbYesNoCancel + vbQuestion, "消費税方式の確認")
        If taxAns = vbCancel Then Exit Sub
        If taxAns = vbYes Then
            taxMode  = "nuki"
            skip税込 = False
        Else
            taxMode  = "komi"
            skip税込 = True   ' 税込方式は税込CSVをスキップ
        End If

        ' 実行前確認（A-3）
        ' スキップ判断（A-4）
        ' Application設定・エラーハンドラ（A-5以降で追記）

        On Error GoTo ErrHandler
        Application.ScreenUpdating = False
        Application.DisplayAlerts  = False
        Application.StatusBar      = "CSV取込を開始します..."

        ' CSV貼り付け処理（A-6で実装）

        GoTo Cleanup

    ErrHandler:
        ' エラーハンドラ（A-7で実装）
        Resume Cleanup

    Cleanup:
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        Application.StatusBar      = False
    End Sub
    ```
  - 完了条件：
    - MsgBox(YesNoCancel)で消費税方式を選択できる
    - キャンセル時は `Exit Sub` で終了する
    - `On Error GoTo ErrHandler` が設定されている
    - `Cleanup:` で Application 状態が必ず復元される

---

### 🟡 Task A-3：実行前確認ダイアログ

**背景：** 誤操作でデータを消してしまうリスクを防ぐ。実行前に「既存データをクリアします」と警告する。

- [ ] **A-3：消費税方式選択の直後に確認ダイアログを追加する**
  - 対象：`vba/CSV自動取込.bas`（A-2の `skip税込 = True` の直後）
  - 実装内容：
    ```vb
        ' --- 実行前確認 ---
        If MsgBox("既存のデータをクリアして取り込みを開始します。" & vbCrLf & _
                  "よろしいですか？", _
                  vbYesNo + vbExclamation, "実行確認") = vbNo Then
            Exit Sub
        End If
    ```
  - 完了条件：「いいえ」でマクロ終了、「はい」で次の処理に進む

---

### 🟡 Task A-4：スキップ判断とシートクリア

**背景：** 現在の旧版は実行冒頭で3シートを全クリアしてしまうため、スキップしたシートのデータも消えてしまう。スキップ判断をクリアより前に行う必要がある。

- [ ] **A-4：三期分スキップ確認と、必要なシートだけクリアするロジックを実装する**
  - 対象：`vba/CSV自動取込.bas`（実行前確認の直後）
  - 実装内容：
    ```vb
        ' --- 三期分CSVのスキップ確認 ---
        Dim san3Ans As Integer
        san3Ans = MsgBox("三期分CSVを取り込みますか？" & vbCrLf & _
                         "（スキップすると既存データをそのまま保持します）", _
                         vbYesNo + vbQuestion, "三期分CSV")
        skip三期 = (san3Ans = vbNo)

        ' --- 必要なシートだけクリア（スキップしたシートは触らない） ---
        wb.Sheets(IDX_当期).Cells.ClearContents
        If Not skip税込 Then wb.Sheets(IDX_税込).Cells.ClearContents
        If Not skip三期 Then wb.Sheets(IDX_三期分).Cells.ClearContents
    ```
  - 完了条件：
    - スキップしたシートは ClearContents されない
    - 税込方式（skip税込=True）のとき Sheet(IDX_税込) はクリアされない

---

### 🔵 Task A-5：CSVを貼付 Function（コア処理）

**背景：** 現在の `Paste Sub` はエラー処理も戻り値もない。Function 化して行数を返すようにし、エラー時の挙動を安全にする。

- [ ] **A-5：`CSVを貼付` Functionをメインとは別に実装する**
  - 対象：`vba/CSV自動取込.bas`（メインSubの外・末尾に追記）
  - 実装内容：
    ```vb
    ' CSVファイルを指定シートのA1から貼り付ける
    ' 戻り値：取り込み行数（成功）/ -1（エラー or スキップ）
    Private Function CSVを貼付(csvPath As String, ws As Worksheet) As Long
        CSVを貼付 = -1
        On Error GoTo PasteError

        ' シートの存在確認
        If ws Is Nothing Then
            MsgBox "貼り付け先シートが見つかりません。", vbCritical, "シートエラー"
            Exit Function
        End If

        ' CSVを開いて貼り付け
        Dim csvWb As Workbook
        Set csvWb = Workbooks.Open(csvPath)

        Dim usedRows As Long
        usedRows = csvWb.Sheets(1).UsedRange.Rows.Count

        ' 空ファイルチェック
        If usedRows <= 1 Then
            csvWb.Close False
            MsgBox "選択されたCSVにデータが含まれていません。スキップします。", _
                   vbExclamation, "空ファイル"
            Exit Function
        End If

        ' 貼り付け実行
        csvWb.Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
        csvWb.Close False
        ThisWorkbook.Activate

        CSVを貼付 = usedRows - 1  ' ヘッダー行を除いたデータ行数
        Exit Function

    PasteError:
        On Error Resume Next
        If Not csvWb Is Nothing Then csvWb.Close False
        On Error GoTo 0

        ' 文字コードエラーの可能性がある場合の案内
        Dim errMsg As String
        errMsg = "CSVの読み込み中にエラーが発生しました。" & vbCrLf & vbCrLf
        If Err.Number = 1004 Or Err.Number = 0 Then
            errMsg = errMsg & "【文字コードの問題の可能性があります】" & vbCrLf & _
                     "対処方法：" & vbCrLf & _
                     "1. ExcelでそのCSVファイルを開く（文字コード：UTF-8を指定）" & vbCrLf & _
                     "2. 「名前を付けて保存」でCSV形式で保存し直す" & vbCrLf & _
                     "3. 再度このマクロを実行してください"
        Else
            errMsg = errMsg & "エラー番号：" & Err.Number & vbCrLf & _
                     "詳細：" & Err.Description
        End If
        MsgBox errMsg, vbCritical, "読み込みエラー"
        CSVを貼付 = -1
    End Function
    ```
  - 完了条件：
    - 成功時に取り込み行数（Long）を返す
    - 空CSV時は警告を出してスキップ（-1を返す）
    - エラー時に csvWb を閉じてからエラーメッセージを表示する
    - 文字コード問題の可能性をユーザーに案内する

---

### 🟡 Task A-6：メインSubの貼り付け処理・進捗・完了メッセージ

**背景：** A-5で作ったFunctionを呼び出して、3種類のCSVを取り込む。各ステップでStatusBarに進捗を表示し、完了後に行数を表示する。

- [ ] **A-6：メインSubの `--- CSV貼り付け処理 ---` 以降を実装する**
  - 対象：`vba/CSV自動取込.bas`（メインSubの `Application.StatusBar = "CSV取込を開始します..."` の直後）
  - 実装内容：
    ```vb
        ' --- ① 当期CSV（必須） ---
        Application.StatusBar = "① 当期CSVを選択してください..."
        Dim path当期 As String
        path当期 = Application.GetOpenFilename("CSVファイル,*.csv", , "① 当期CSVを選択（税抜）")
        If path当期 = "False" Then
            If MsgBox("当期CSVの選択がキャンセルされました。" & vbCrLf & _
                      "処理を中止しますか？", vbYesNo + vbQuestion, "キャンセル確認") = vbYes Then
                GoTo Cleanup
            End If
        Else
            Application.StatusBar = "① 当期CSVを貼り付け中..."
            rows当期 = CSVを貼付(path当期, wb.Sheets(IDX_当期))
        End If

        ' --- ② 税込CSV（税抜方式のみ） ---
        If Not skip税込 Then
            Application.StatusBar = "② 税込CSVを選択してください..."
            Dim path税込 As String
            path税込 = Application.GetOpenFilename("CSVファイル,*.csv", , "② 税込CSVを選択")
            If path税込 = "False" Then
                MsgBox "税込CSVをスキップしました。税込データ専用シートは変更されません。", _
                       vbInformation, "スキップ"
            Else
                Application.StatusBar = "② 税込CSVを貼り付け中..."
                rows税込 = CSVを貼付(path税込, wb.Sheets(IDX_税込))
            End If
        End If

        ' --- ③ 三期分CSV（任意） ---
        If Not skip三期 Then
            Application.StatusBar = "③ 三期分CSVを選択してください..."
            Dim path三期 As String
            path三期 = Application.GetOpenFilename("CSVファイル,*.csv", , "③ 三期分CSVを選択")
            If path三期 = "False" Then
                MsgBox "三期分CSVをスキップしました。三期分シートは変更されません。", _
                       vbInformation, "スキップ"
            Else
                Application.StatusBar = "③ 三期分CSVを貼り付け中..."
                rows三期 = CSVを貼付(path三期, wb.Sheets(IDX_三期分))
            End If
        End If

        ' --- 完了メッセージ ---
        Application.StatusBar = "取り込み完了！"
        Dim msg As String
        msg = "CSV取り込みが完了しました。" & vbCrLf & vbCrLf
        If rows当期 > 0  Then msg = msg & "① 当期：  " & rows当期 & " 行" & vbCrLf
        If rows税込 > 0  Then msg = msg & "② 税込：  " & rows税込 & " 行" & vbCrLf
        If rows三期 > 0  Then msg = msg & "③ 三期分：" & rows三期 & " 行" & vbCrLf
        If taxMode = "komi" Then
            msg = msg & vbCrLf & "※ 税込方式：税込データは数式で自動連動"
        End If
        MsgBox msg, vbInformation, "完了"
    ```
  - 完了条件：
    - 当期CSVキャンセル時に「中止しますか？」を確認する
    - StatusBar で各ステップの進捗を表示する
    - 完了MsgBox に取り込み行数を表示する

---

### 🟡 Task A-7：ErrHandlerの実装

**背景：** 予期せぬエラー発生時に ScreenUpdating 等が False のまま残るとExcel全体が操作不能になる。必ず Cleanup に通す。

- [ ] **A-7：`ErrHandler:` ラベルを実装する**
  - 対象：`vba/CSV自動取込.bas`（`GoTo Cleanup` の直後）
  - 実装内容：
    ```vb
    ErrHandler:
        MsgBox "予期せぬエラーが発生しました。" & vbCrLf & _
               "エラー番号：" & Err.Number & vbCrLf & _
               "詳細：" & Err.Description, vbCritical, "エラー"
        Resume Cleanup
    ```
  - 完了条件：エラー発生時に `Cleanup:` へジャンプして Application 状態を復元する

---

## 【実装セッションB】デモUI追加画面

**指示：B-1a → B-1b → ... → B-3 の順に実装してください。対象は `demo/index.html` 1ファイルのみ。外部ライブラリ禁止。既存CSSクラスを流用してください。**

---

### 🟡 Task B-1a：SCREEN 1b「1表紙」追加

**背景：** クライアント向けデモに表紙画面がない。会社名・会計期間・担当者の確認画面を追加する。

- [ ] **B-1a：`screen-1b` を追加する（`screen-0` の直後、`screen-1` の直前に挿入）**
  - 会社名・会計期間（YYYY年MM月〜YYYY年MM月）・担当者名の入力フォーム
  - 「確認して次へ」ボタン → `goSheet('1c')`
  - タイトルバーのファイル名：`1表紙.xlsx`
  - 完了条件：画面が表示され、次へボタンで遷移できる

---

### 🟡 Task B-1b：SCREEN 1c「2月別推移グラフ」追加

- [ ] **B-1b：`screen-1c` を追加する**
  - 売上・粗利の月別棒グラフ（SVGかCSSバー。Chart.js禁止）
  - 4月〜3月の12ヶ月分モックデータ
  - 「レポート生成」ボタン + プログレスバーアニメーション
  - 「次へ」ボタン → `goSheet('1d')`
  - タイトルバーのファイル名：`2月別推移グラフ.xlsx`
  - 完了条件：グラフ表示・生成ボタン動作・遷移できる

---

### 🟡 Task B-1c：SCREEN 1d「3売上三期比較グラフ」追加

- [ ] **B-1c：`screen-1d` を追加する**
  - 前々期・前期・当期の3本並び棒グラフ（SVGかCSSバー）
  - 「レポート生成」ボタン + プログレスバー
  - 「次へ」ボタン → `goSheet('1')`（既存の4月次収益分析表へ）
  - タイトルバーのファイル名：`3売上三期比較グラフ.xlsx`
  - 完了条件：3期分グラフ表示・遷移できる

---

### 🟡 Task B-1d：SCREEN 1e「5三期比較決算図表」追加

- [ ] **B-1d：`screen-1e` を追加する**
  - 3期分のPL（売上・粗利・営業利益）とBS（資産・負債・純資産）を横並びテーブル
  - 「レポート生成」ボタン + プログレスバー
  - 「次へ」ボタン → `goSheet('1f')`
  - タイトルバーのファイル名：`5三期比較決算図表.xlsm`
  - 完了条件：テーブル表示・遷移できる

---

### 🟡 Task B-1e：SCREEN 1f「6経営計画・予実グラフ」追加

- [ ] **B-1e：`screen-1f` を追加する**
  - 計画vs実績の折れ線グラフ（SVGで実装。X軸=月、Y軸=金額）
  - 「レポート生成」ボタン + プログレスバー
  - 「次へ」ボタン → `goSheet('1g')`
  - タイトルバーのファイル名：`6経営計画･予実グラフ.xlsm`
  - 完了条件：折れ線グラフ表示・遷移できる

---

### 🟡 Task B-1f：SCREEN 1g「10キャッシュ推移表」追加

- [ ] **B-1f：`screen-1g` を追加する**
  - 月次キャッシュ残高の推移グラフ（SVGかCSSバー）
  - 運転資金倍率表（月×倍率のモックテーブル）
  - 「レポート生成」ボタン + プログレスバー
  - 「次へ」ボタン → 既存の `screen-2`（11CF計算書）へ
  - タイトルバーのファイル名：`10ｷｬｯｼｭ推移表～運転資金倍率表.xlsm`
  - 完了条件：グラフ・テーブル表示・遷移できる

---

### 🟢 Task B-2：画面遷移・タブ・STEPドット 全体改修

**依存：B-1a〜B-1f が全て完了してから実施**

- [ ] **B-2：`goSheet()` と シートタブを10画面対応に改修する**
  - シートタブを以下の10タブに更新：
    CSV取込 / 1表紙 / 2月別推移 / 3三期比較 / 4月次分析 / 5三期決算 / 6経営計画 / 10CF推移 / 11CF計算書 / 12決算書
  - STEPドットを10個に変更
  - `goSheet()` の `fnames` マップを全画面のファイル名に対応させる
  - 完了条件：全タブが機能し、STEPドットが正しく表示され、ファイル名が切り替わる

---

### 🟢 Task B-3：完了画面の改修

**依存：B-2 が完了してから実施**

- [ ] **B-3：`screen-4` の完了リストを全ファイル対応に更新し、PDFボタンを追加する**
  - 完了リスト：1表紙〜12決算書図表の全ファイルを列挙
  - 「PDFとして保存」ボタン（`window.print()` を呼び出す。スタイルは `btn-blue`）
  - 完了条件：全ファイルが完了リストに表示され、PDFボタンが存在する

---

## 判断ルール集（聞かずに自己判断してよいこと）

| 場面 | 判断ルール |
|------|-----------|
| VBAのシートインデックス | Const定数（IDX_当期/IDX_税込/IDX_三期分）を使う。直書き禁止 |
| VBAのエラー番号判定 | 正確な判定が困難な場合は「〜の可能性があります」という案内文で実装 |
| VBAのコメント | すべて日本語 |
| デモUIのグラフ | Chart.js等の外部ライブラリ禁止。SVGインライン or CSSバーで実装 |
| デモUIのモックデータ | 財務らしい数字（売上1000〜5000万円台）であれば何でもよい |
| デモUIのスタイル | 既存のCSSクラス（`btn-green` / `section` / `section-title` 等）を流用 |
| 画面IDの命名 | `screen-{ファイル番号}` パターン（例：`screen-1b`） |

## エラーハンドリング方針

| 状況 | 対応方針 |
|------|---------|
| ファイル選択キャンセル | 当期CSVのみ「中止しますか？」を確認。②③はスキップ扱い |
| シートが見つからない | エラーメッセージを表示して -1 を返す |
| 空CSVファイル | 警告を出してスキップ（-1を返す） |
| 予期せぬVBAエラー | ErrHandler → Resume Cleanup でApplication状態を必ず復元 |
| 文字コード問題 | エラー時に「Shift-JIS以外の可能性あり」＋対処手順3ステップを案内 |

---

## 🐛 エラー・バグ管理

### 未対応
（実装中に発見したバグをここに追記する）

### 対応済み
（修正完了したものをここに移動する）

---

## 引き継ぎメモ

### 2026-04-15 CLAUDE.md 全面改修

- leopalaceプロジェクトのCLAUDE.mdスタイルに合わせて全面書き直し
- 実装セッションA（VBAマクロ A-1〜A-7）と実装セッションB（デモUI B-1a〜B-3）に分離
- 各タスクに背景・実装コード・完了条件を記載済み
- `vba/CSV自動取込.bas` は旧版（59行）のまま。次の実装セッションでA-1から全面書き直し
- 次のセッションは「実装セッションAを実行してください」と伝えるだけでOK
