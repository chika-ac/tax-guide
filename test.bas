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

    Application.ScreenUpdating = False

    ' 選択範囲をコピー
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' 一時シートに貼り付け
    Set tmpSheet = Worksheets.Add
    tmpSheet.Paste

    ' 貼り付けた図オブジェクトを直接エクスポート
    Set shp = tmpSheet.Shapes(1)
    shp.Export Filename:=savePath, FilterName:="PNG"

    ' 一時シートを削除
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    MsgBox "保存しました。" & vbCrLf & savePath, vbInformation

End Sub
