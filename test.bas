Sub RangeToImage()

    ' =============================================
    ' 設定：必要に応じてここを変更してください
    ' =============================================
    Const SAVE_FOLDER As String = "C:\Users\YourName\Pictures\"  ' 保存先フォルダ（末尾に\をつける）
    Const FILE_NAME   As String = "output.png"                    ' ファイル名
    ' =============================================

    Dim rng         As Range
    Dim tmpSheet    As Worksheet
    Dim shp         As Shape
    Dim chtObj      As ChartObject
    Dim cht         As Chart
    Dim savePath    As String

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

    ' 一時シートを追加して貼り付け
    Application.ScreenUpdating = False
    Set tmpSheet = Worksheets.Add

    tmpSheet.Paste

    ' 貼り付けた図をグラフ経由でエクスポート
    Set shp = tmpSheet.Shapes(tmpSheet.Shapes.Count)
    shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    Set chtObj = tmpSheet.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    Set cht = chtObj.Chart
    cht.Paste
    cht.Export Filename:=savePath, FilterName:="PNG"

    ' 一時シートを削除
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "保存しました。" & vbCrLf & savePath, vbInformation

End Sub
