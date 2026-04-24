Attribute VB_Name = "MaskPII"
Option Explicit

' ================================================================
' 個人情報マスクマクロ  ver.1.0
' 対象：職番（7〜8桁）、申請番号（12桁）、
'       氏名（さん/様/さま/サマ付き）、年収・金額（単位付き）
' ================================================================

Private Const MASK_TOKEN As String = "[マスク]"

' ----------------------------------------------------------------
' 【メイン①】開いているシートをコピーしてマスク版を別シートに作成
'   手順: CSVをExcelで開く → このマクロ実行 → "_masked"シートが追加される
' ----------------------------------------------------------------
Sub MaskToNewSheet()
    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    ' シートをコピーして末尾に追加
    srcWs.Copy After:=Sheets(Sheets.Count)
    Dim newWs As Worksheet
    Set newWs = ActiveSheet
    newWs.Name = Left(srcWs.Name, 27) & "_masked"

    Dim lastRow As Long, lastCol As Long
    lastRow = newWs.Cells(newWs.Rows.Count, 1).End(xlUp).Row
    lastCol = newWs.Cells(1, newWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "データがありません（2行目以降が処理対象です）。"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim changedCount As Long
    changedCount = 0
    Dim i As Long, j As Long

    ' 1行目はヘッダーなのでスキップ
    For i = 2 To lastRow
        For j = 1 To lastCol
            Dim cellVal As String
            cellVal = CStr(newWs.Cells(i, j).Value)
            If Len(cellVal) > 0 Then
                Dim maskedVal As String
                maskedVal = MaskText(cellVal)
                If maskedVal <> cellVal Then
                    newWs.Cells(i, j).Value = maskedVal
                    changedCount = changedCount + 1
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "完了！シート「" & newWs.Name & "」を作成しました。" & vbCrLf & _
           "変更セル数: " & changedCount & " 個"
End Sub

' ----------------------------------------------------------------
' 【メイン②】アクティブシートを直接マスク（上書き）
'   ※元に戻せないので先にバックアップを取ること
' ----------------------------------------------------------------
Sub MaskActiveSheet()
    If MsgBox("アクティブシートを直接上書きマスクします。" & vbCrLf & _
              "元に戻せません。よろしいですか？", _
              vbYesNo + vbExclamation, "確認") = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "データがありません（2行目以降が処理対象です）。"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim changedCount As Long
    changedCount = 0
    Dim i As Long, j As Long

    For i = 2 To lastRow
        For j = 1 To lastCol
            Dim cellVal As String
            cellVal = CStr(ws.Cells(i, j).Value)
            If Len(cellVal) > 0 Then
                Dim maskedVal As String
                maskedVal = MaskText(cellVal)
                If maskedVal <> cellVal Then
                    ws.Cells(i, j).Value = maskedVal
                    changedCount = changedCount + 1
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "完了！変更セル数: " & changedCount & " 個"
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
    ' グループ(0)=名前, グループ(1)=敬称
    re.Pattern = "([一-龥ぁ-ゖァ-ヴー]{1,12}(?:\s*[一-龥ぁ-ゖァ-ヴー]{1,12})?)\s*(さん|様|さま|サマ)"

    Dim matches As Object
    Set matches = re.Execute(text)
    If matches.Count = 0 Then
        MaskNameWithHonorific = text
        Exit Function
    End If

    ' 後ろから置換して文字列位置のズレを防ぐ
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
    ' 万円 を 円 より先にマッチさせるため万円を前に配置
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
    ' \ または ¥ に続く数字をマスク（記号自体は残す）
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
