Sub RangeToImage()

    ' =============================================
    ' 設定：必要に応じてここを変更してください
    ' =============================================
    Const SAVE_FOLDER As String = "C:\Users\YourName\Pictures\"  ' 保存先フォルダ（末尾に\をつける）
    Const FILE_NAME   As String = "output.png"                    ' ファイル名
    ' =============================================

    Dim rng     As Range
    Dim cht     As Chart
    Dim chtObj  As ChartObject
    Dim savePath As String

    ' 選択範囲の確認
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    savePath = SAVE_FOLDER & FILE_NAME

    ' 保存先フォルダの存在確認
    If Dir(SAVE_FOLDER, vbDirectory) = "" Then
        MsgBox "保存先フォルダが見つかりません。" & vbCrLf & SAVE_FOLDER, vbExclamation
        Exit Sub
    End If

    ' 選択範囲をコピー
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' 一時グラフを作成して貼り付け
    Set chtObj = ActiveSheet.ChartObjects.Add(0, 0, rng.Width, rng.Height)
    Set cht = chtObj.Chart

    With cht
        .Paste
        .Export Filename:=savePath, FilterName:="PNG"
    End With

    ' 一時グラフを削除
    chtObj.Delete

    MsgBox "保存しました。" & vbCrLf & savePath, vbInformation

End Sub
