Attribute VB_Name = "modSameCheck"
Option Explicit

'==============================================================
' SameCheck - 重複データ抽出・ハイライトマクロ
'
' 指定した1つのCSVファイルを読み込み、全列が完全一致する行を
' 「重複」として検出します。
' ・元データシートでは、重複している行すべてに色を付けます
' ・「重複一覧」シートには、重複しているデータの件数・「データ」
'   シート上の該当行番号・元の行内容を一覧でまとめます
'
' 作者: 天電輝合同会社 (TENDENKO)
' ライセンス: MIT License
'
' ---- 使い方 ----
' 1. リボンまたはAlt+F8から「Main_重複チェック実行」を実行
' 2. ファイル選択ダイアログで、チェックしたいCSVファイルを選ぶ
' 3. 「データ」シートに全件が読み込まれ、重複行には色が付く
'    「重複一覧」シートに重複データのサマリーが出力される
'
' ---- カスタマイズポイント（下の設定セクションを書き換えるだけ） ----
' ・INPUT_ENCODING    : 読み込むCSVの文字コード
' ・HAS_HEADER        : 1行目がヘッダー行かどうか
' ・DATA_SHEET_NAME   : 元データを書き出すシート名
' ・DUPLICATE_SHEET_NAME : 重複一覧を書き出すシート名
'
' ---- 判定方法 ----
' 行のすべての列（全文字列）が完全に一致する場合のみ「重複」と
' 判定します。一部の列だけ一致するケースは対象外です。
' CSVのダブルクォート囲みや、区切り文字自体を含む値には対応して
' いません（単純な区切り文字分割）。
'==============================================================

'---------------- 設定セクション（ここだけ書き換えればOK） ----------------

' 文字コード。"shift-jis" か "utf-8" を指定（大文字小文字どちらでも可）
Private Const INPUT_ENCODING As String = "shift-jis"

' CSVの1行目がヘッダー（項目名）かどうか。
' True の場合、1行目は重複判定の対象から除外します。
Private Const HAS_HEADER As Boolean = True

' 区切り文字（通常は "," のままでOK。タブ区切りなら vbTab）
Private Const DELIMITER As String = ","

' 元データを書き出すシート名
Private Const DATA_SHEET_NAME As String = "データ"

' 重複一覧を書き出すシート名
Private Const DUPLICATE_SHEET_NAME As String = "重複一覧"

' 重複グループごとに、GetColorPalette() のパレットから順番に色を割り当てます。
' グループ数がパレットの色数より多い場合は、最初の色に戻って繰り返し使用します。

'---------------------------------------------------------------------

'==============================================================
' Main_重複チェック実行
'   エントリーポイント。ボタンやAlt+F8から呼び出す。
'==============================================================
Sub Main_重複チェック実行()

    Dim targetFile As String

    On Error GoTo ErrHandler

    ' --- ファイル選択 ---
    targetFile = PickCsvFile()
    If targetFile = "" Then
        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbInformation
        Exit Sub
    End If

    ' --- CSV読み込み ---
    Dim lines() As String
    Dim lineCount As Long
    lineCount = ReadTextFileLines(targetFile, INPUT_ENCODING, lines)

    If lineCount = 0 Then
        MsgBox "ファイルの読み込みに失敗したか、空のファイルです。" & vbCrLf & _
               "ファイル: " & targetFile, vbExclamation
        Exit Sub
    End If

    Dim startRow As Long
    startRow = 0
    If HAS_HEADER Then startRow = 1

    If startRow >= lineCount Then
        MsgBox "ヘッダー行以外にデータがありませんでした。", vbExclamation
        Exit Sub
    End If

    ' --- 重複判定（キー = 行全体のテキスト） ---
    ' rowMap: キー(行内容) → その行が「データ」シート上で出現する行番号のCollection
    ' 「データ」シート上の行番号をここで一括計算しておくことで、
    ' 「重複一覧」シート側でも「何行目と何行目が重複しているか」を示せるようにする。
    Dim rowMap As Object
    Set rowMap = CreateObject("Scripting.Dictionary")

    Dim headerOffset As Long
    headerOffset = IIf(HAS_HEADER, 1, 0)

    Dim r As Long
    Dim key As String
    Dim dataSheetRow As Long
    Dim rowList As Collection

    For r = startRow To lineCount - 1
        key = lines(r)
        dataSheetRow = (r - startRow) + 1 + headerOffset ' データシート上での行番号（ヘッダー行の分を考慮）

        If rowMap.Exists(key) Then
            rowMap(key).Add dataSheetRow
        Else
            ' 毎回必ず新しいCollectionを作る（"Dim ... As New" をループ内で使うと
            ' 全キーで同じインスタンスを使い回してしまうため、明示的にSetする）
            Set rowList = New Collection
            rowList.Add dataSheetRow
            rowMap.Add key, rowList
        End If
    Next r

    ' --- 重複グループごとに、見分けやすいよう別々の色を割り当てる ---
    ' （「データ」シートと「重複一覧」シートで同じグループには同じ色を使うので、
    '   どの行とどの行がセットになっているか、色でも対応がわかる）
    Dim colorMap As Object
    Set colorMap = CreateObject("Scripting.Dictionary")

    Dim palette As Variant
    palette = GetColorPalette()

    Dim colorIndex As Long
    colorIndex = 0

    Dim k As Variant
    For Each k In rowMap.Keys
        If rowMap(k).Count > 1 Then
            colorMap.Add k, palette(colorIndex Mod (UBound(palette) - LBound(palette) + 1))
            colorIndex = colorIndex + 1
        End If
    Next k

    ' --- 「データ」シートへ書き出し（重複行はグループごとに色分け） ---
    Dim duplicateRowCount As Long
    duplicateRowCount = WriteDataSheet(lines, lineCount, startRow, rowMap, colorMap)

    ' --- 「重複一覧」シートへ書き出し ---
    Dim duplicateKindCount As Long
    duplicateKindCount = WriteDuplicateSheet(lines, lineCount, startRow, rowMap, colorMap)

    Application.StatusBar = False

    ' --- 完了メッセージ ---
    Dim msg As String
    msg = "重複チェックが完了しました。" & vbCrLf & _
          "対象データ行数: " & (lineCount - startRow) & vbCrLf & vbCrLf & _
          "結果は「" & DATA_SHEET_NAME & "」シートに出力しました。"

    If duplicateKindCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "【重複あり】" & duplicateKindCount & " 種類・計 " & duplicateRowCount & " 行の重複を検出しました。" & vbCrLf & _
              "詳細は「" & DUPLICATE_SHEET_NAME & "」シートをご確認ください。"
        MsgBox msg, vbExclamation, "完了（重複あり）"
    Else
        msg = msg & vbCrLf & vbCrLf & "重複データは見つかりませんでした。"
        MsgBox msg, vbInformation, "完了"
    End If

    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー"
End Sub

'==============================================================
' Main_結果シート削除
'   「データ」「重複一覧」シートを削除する。
'   配布・再テスト前のクリーンアップ用のエントリーポイント。
'==============================================================
Sub Main_結果シート削除()

    Dim targets As Variant
    targets = Array(DATA_SHEET_NAME, DUPLICATE_SHEET_NAME)

    Dim existCount As Long
    existCount = 0

    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        If SheetExists(CStr(targets(i))) Then
            existCount = existCount + 1
        End If
    Next i

    If existCount = 0 Then
        MsgBox "「" & DATA_SHEET_NAME & "」「" & DUPLICATE_SHEET_NAME & "」シートはどちらも存在しません。", _
               vbInformation, "削除不要"
        Exit Sub
    End If

    ' 削除すると全シートがなくなってしまう場合は中止（Excelは最低1シート必要）
    If ThisWorkbook.Worksheets.Count - existCount < 1 Then
        MsgBox "削除すると全てのシートがなくなってしまうため中止しました。" & vbCrLf & _
               "操作用のシートなどを別に1つ残してから実行してください。", vbCritical, "中止"
        Exit Sub
    End If

    If MsgBox("「" & DATA_SHEET_NAME & "」「" & DUPLICATE_SHEET_NAME & "」シートを削除します。" & vbCrLf & _
              "元に戻せません。よろしいですか？", vbYesNo + vbExclamation, "確認") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False

    For i = LBound(targets) To UBound(targets)
        If SheetExists(CStr(targets(i))) Then
            ThisWorkbook.Worksheets(CStr(targets(i))).Delete
        End If
    Next i

    Application.DisplayAlerts = True

    MsgBox "削除が完了しました。", vbInformation, "完了"

End Sub

'==============================================================
' 指定シートが存在するかどうかを判定する
'==============================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

'==============================================================
' ファイル選択ダイアログを表示し、選択されたCSVパスを返す
'==============================================================
Private Function PickCsvFile() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.Title = "重複をチェックしたいCSVファイルを選択してください"
    fd.AllowMultiSelect = False

    fd.Filters.Clear
    fd.Filters.Add "CSVファイル", "*.csv"

    ' 初期表示フォルダ：このブックが置かれているフォルダを開く
    If ThisWorkbook.Path <> "" Then
        fd.InitialFileName = ThisWorkbook.Path & "\"
    End If

    If fd.Show = -1 Then
        PickCsvFile = fd.SelectedItems(1)
    Else
        PickCsvFile = ""
    End If

End Function

'==============================================================
' テキストファイルを指定エンコーディングで読み込み、行配列を返す
' 戻り値: 行数
'==============================================================
Private Function ReadTextFileLines(ByVal filePath As String, ByVal encoding As String, ByRef lines() As String) As Long

    On Error GoTo ErrHandler

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2 ' adTypeText
    stm.Charset = NormalizeCharset(encoding)
    stm.Open
    stm.LoadFromFile filePath

    Dim content As String
    content = stm.ReadText
    stm.Close

    ' 改行コードの統一（CRLF / CR / LF すべてに対応）
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    ' 末尾の空行を除去してから分割
    Do While Len(content) > 0 And Right(content, 1) = vbLf
        content = Left(content, Len(content) - 1)
    Loop

    If Len(content) = 0 Then
        ReadTextFileLines = 0
        Exit Function
    End If

    lines = Split(content, vbLf)
    ReadTextFileLines = UBound(lines) - LBound(lines) + 1

    Exit Function

ErrHandler:
    ReadTextFileLines = 0
End Function

'==============================================================
' "shift-jis" / "utf-8" などの表記ゆれをADODBが認識する名称に変換
'==============================================================
Private Function NormalizeCharset(ByVal encoding As String) As String

    Dim e As String
    e = LCase(Trim(encoding))

    Select Case e
        Case "sjis", "shift-jis", "shift_jis", "shiftjis"
            NormalizeCharset = "shift-jis"
        Case "utf8", "utf-8"
            NormalizeCharset = "utf-8"
        Case "utf8bom", "utf-8-bom"
            NormalizeCharset = "utf-8"
        Case Else
            NormalizeCharset = e
    End Select

End Function

'==============================================================
' 「データ」シートに全行を書き出し、重複している行をハイライトする
' 戻り値: ハイライトした行数（重複している行の合計件数）
'==============================================================
Private Function WriteDataSheet(ByRef lines() As String, ByVal lineCount As Long, _
                                 ByVal startRow As Long, ByRef rowMap As Object, _
                                 ByRef colorMap As Object) As Long

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(DATA_SHEET_NAME)
    ws.Cells.Clear

    Dim outRow As Long
    outRow = 1
    Dim highlightCount As Long
    highlightCount = 0

    ' ヘッダー行
    If HAS_HEADER Then
        Dim headerFields() As String
        headerFields = Split(lines(0), DELIMITER)
        WriteRowToSheet ws, outRow, headerFields
        ws.Rows(outRow).Font.Bold = True
        ws.Rows(outRow).Interior.Color = RGB(221, 235, 247)
        outRow = outRow + 1
    End If

    Dim r As Long
    For r = startRow To lineCount - 1
        Dim rowFields() As String
        rowFields = Split(lines(r), DELIMITER)
        WriteRowToSheet ws, outRow, rowFields

        If rowMap(lines(r)).Count > 1 Then
            ws.Rows(outRow).Interior.Color = colorMap(lines(r))
            highlightCount = highlightCount + 1
        End If

        outRow = outRow + 1
    Next r

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A1").Select

    WriteDataSheet = highlightCount

End Function

'==============================================================
' 1行分のデータ（配列）を、指定行に書き込む
'==============================================================
Private Sub WriteRowToSheet(ByRef ws As Worksheet, ByVal rowIndex As Long, ByRef fields() As String)

    Dim c As Long
    For c = LBound(fields) To UBound(fields)
        ws.Cells(rowIndex, c - LBound(fields) + 1).Value = fields(c)
    Next c

End Sub

'==============================================================
' 「重複一覧」シートに、重複しているデータのサマリーを書き出す
' 戻り値: 重複している「種類」の数（同じ内容のグループ数）
'==============================================================
Private Function WriteDuplicateSheet(ByRef lines() As String, ByVal lineCount As Long, _
                                      ByVal startRow As Long, ByRef rowMap As Object, _
                                      ByRef colorMap As Object) As Long

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(DUPLICATE_SHEET_NAME)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "実行日時"
    ws.Cells(1, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(3, 1).Value = "件数"
    ws.Cells(3, 2).Value = "「" & DATA_SHEET_NAME & "」シートの該当行"
    ws.Cells(3, 3).Value = "元データの行内容"
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 2).Font.Bold = True
    ws.Cells(3, 3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(221, 235, 247)

    Dim outRow As Long
    outRow = 4
    Dim kindCount As Long
    kindCount = 0

    ' 出現順を保つため、元データを再走査しながら「初めて見つけた重複キー」だけを出力する
    Dim reportedKeys As Object
    Set reportedKeys = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = startRow To lineCount - 1
        Dim key As String
        key = lines(r)

        If rowMap(key).Count > 1 And Not reportedKeys.Exists(key) Then
            ws.Cells(outRow, 1).Value = rowMap(key).Count
            ws.Cells(outRow, 2).Value = JoinRowNumbers(rowMap(key))
            ws.Cells(outRow, 3).Value = key
            ws.Rows(outRow).Interior.Color = colorMap(key)
            reportedKeys.Add key, True
            kindCount = kindCount + 1
            outRow = outRow + 1
        End If
    Next r

    If kindCount = 0 Then
        ws.Cells(outRow, 1).Value = "重複データは見つかりませんでした。"
    End If

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 120

    WriteDuplicateSheet = kindCount

End Function

'==============================================================
' 行番号のCollectionを、"4, 5, 12" のようなカンマ区切り文字列に変換する
'==============================================================
Private Function JoinRowNumbers(ByRef rowNumbers As Collection) As String

    Dim result As String
    result = ""

    Dim i As Long
    For i = 1 To rowNumbers.Count
        If i > 1 Then result = result & ", "
        result = result & rowNumbers(i)
    Next i

    JoinRowNumbers = result

End Function

'==============================================================
' 重複グループの色分けに使うパレット（8色）を返す。
' グループ数がこれより多い場合は、先頭の色に戻って繰り返し使う。
'==============================================================
Private Function GetColorPalette() As Variant
    ' 赤、黄、緑、青、紫、ピンク、ミント、オレンジ の順
    GetColorPalette = Array( _
        RGB(255, 199, 206), _
        RGB(255, 230, 153), _
        RGB(198, 239, 206), _
        RGB(189, 215, 238), _
        RGB(221, 204, 255), _
        RGB(255, 204, 229), _
        RGB(204, 255, 229), _
        RGB(255, 214, 153) _
    )
End Function

'==============================================================
' シートが存在すれば取得、なければ新規作成して返す
'==============================================================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws

End Function
