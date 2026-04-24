Attribute VB_Name = "MaskPII"
Option Explicit

' ================================================================
' 個人情報マスクマクロ  ver.2.0
' 動作：「問い合わせ履歴」シートのデータを「マスク済み」シートに
'        マスクして転記する
' ボタンから呼び出す関数： RunMask
' ================================================================

Private Const MASK_TOKEN    As String = "[マスク]"
Private Const SRC_SHEET     As String = "問い合わせ履歴"
Private Const DST_SHEET     As String = "マスク済み"

' ----------------------------------------------------------------
' ボタンに割り当てる関数（エントリーポイント）
' ----------------------------------------------------------------
Public Sub RunMask()

    ' ── 元シートの確認 ──────────────────────────────────────────
    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets(SRC_SHEET)
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "シート「" & SRC_SHEET & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    ' ── 転記先シートを用意（なければ作成、あれば中身をクリア） ──
    Dim dstWs As Worksheet
    On Error Resume Next
    Set dstWs = ThisWorkbook.Sheets(DST_SHEET)
    On Error GoTo 0

    If dstWs Is Nothing Then
        ' 末尾に新規作成
        Set dstWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dstWs.Name = DST_SHEET
    Else
        dstWs.Cells.Clear
    End If

    ' ── データ範囲を取得 ────────────────────────────────────────
    Dim lastRow As Long, lastCol As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Then
        MsgBox "「" & SRC_SHEET & "」にデータがありません。"
        Exit Sub
    End If

    ' ── 元データをそのまま転記 ──────────────────────────────────
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol)).Copy
    dstWs.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' ── 2行目以降（ヘッダー除く）をマスク ──────────────────────
    Dim changedCount As Long
    changedCount = 0
    Dim i As Long, j As Long

    For i = 2 To lastRow
        For j = 1 To lastCol
            Dim cellVal As String
            cellVal = CStr(dstWs.Cells(i, j).Value)
            If Len(cellVal) > 0 Then
                Dim maskedVal As String
                maskedVal = MaskText(cellVal)
                If maskedVal <> cellVal Then
                    dstWs.Cells(i, j).Value = maskedVal
                    changedCount = changedCount + 1
                End If
            End If
        Next j
    Next i

    ' ── 「マスク済み」シートをアクティブにして完了通知 ───────────
    dstWs.Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "完了！「" & DST_SHEET & "」シートに転記しました。" & vbCrLf & _
           "処理行数: " & (lastRow - 1) & " 行　／　変更セル: " & changedCount & " 個"

End Sub

' ================================================================
' テキストマスク処理（コア）
' ================================================================
Public Function MaskText(text As String) As String
    Dim result As String
    result = text

    ' ① 申請番号（12桁）を先に処理（長い方を優先）
    result = ReplaceAll(result, "\b\d{12}\b", MASK_TOKEN)

    ' ② 職番（7〜8桁）
    result = ReplaceAll(result, "\b\d{7,8}\b", MASK_TOKEN)

    ' ③ 氏名（敬称付き）：名前だけマスク、敬称（さん/様/さま/サマ）は残す
    result = MaskNameWithHonorific(result)

    ' ④ 金額（単位付き）：数字だけマスク、単位（円/万円/えん/万）は残す
    result = MaskMoneyWithUnit(result)

    ' ⑤ 金額（\・¥記号付き）：記号は残して数字をマスク
    result = MaskMoneyWithYenSign(result)

    MaskText = result
End Function

' ----------------------------------------------------------------
' シンプルな正規表現全置換
' ----------------------------------------------------------------
Private Function ReplaceAll(text As String, pattern As String, _
                            replacement As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = pattern
    ReplaceAll = re.Replace(text, replacement)
End Function

' ----------------------------------------------------------------
' 氏名マスク：名前部分 → [マスク]、敬称はそのまま
' ----------------------------------------------------------------
Private Function MaskNameWithHonorific(text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([一-龥ぁ-ゖァ-ヴー]{1,12}(?:\s*[一-龥ぁ-ゖァ-ヴー]{1,12})?)\s*(さん|様|さま|サマ)"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches.Count = 0 Then
        MaskNameWithHonorific = text
        Exit Function
    End If

    Dim result As String
    result = text
    Dim i As Integer
    For i = matches.Count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches(i)
        result = Left(result, m.FirstIndex) & _
                 MASK_TOKEN & m.SubMatches(1) & _
                 Mid(result, m.FirstIndex + m.Length + 1)
    Next i

    MaskNameWithHonorific = result
End Function

' ----------------------------------------------------------------
' 金額マスク（単位付き）：数字 → [マスク]、単位は残す
' ----------------------------------------------------------------
Private Function MaskMoneyWithUnit(text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(\d{1,3}(?:,\d{3})+|\d+)\s*(万円|えん|円|万)"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches.Count = 0 Then
        MaskMoneyWithUnit = text
        Exit Function
    End If

    Dim result As String
    result = text
    Dim i As Integer
    For i = matches.Count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches(i)
        result = Left(result, m.FirstIndex) & _
                 MASK_TOKEN & m.SubMatches(1) & _
                 Mid(result, m.FirstIndex + m.Length + 1)
    Next i

    MaskMoneyWithUnit = result
End Function

' ----------------------------------------------------------------
' 金額マスク（\・¥記号付き）：記号は残して数字をマスク
' ----------------------------------------------------------------
Private Function MaskMoneyWithYenSign(text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([\\¥]\s*)(\d{1,3}(?:,\d{3})+|\d+)"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches.Count = 0 Then
        MaskMoneyWithYenSign = text
        Exit Function
    End If

    Dim result As String
    result = text
    Dim i As Integer
    For i = matches.Count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches(i)
        result = Left(result, m.FirstIndex) & _
                 m.SubMatches(0) & MASK_TOKEN & _
                 Mid(result, m.FirstIndex + m.Length + 1)
    Next i

    MaskMoneyWithYenSign = result
End Function
