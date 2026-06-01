Attribute VB_Name = "LatexInserter"
Option Explicit

' ============================================================
' LaTeX Equation Inserter for PowerPoint  (PPAM版 v1.0)
' 完全オフライン動作 - latex-renderer.exe を使用
' TeX Live・インターネット接続不要
' ============================================================

Private Const ADDIN_NAME  As String = "LaTeXInserter"
Private Const RENDERER_EX As String = "latex-renderer.exe"
Private Const DEFAULT_SCALE As String = "4"   ' 4× → ~300 DPI

' ---- メインエントリポイント：数式を挿入 ----
Public Sub InsertLatex()
    Dim latex As String
    latex = InputBox( _
        "LaTeX 数式を入力してください:" & vbCrLf & vbCrLf & _
        "例)  \frac{a}{b}            分数" & vbCrLf & _
        "     \int_0^\infty e^{-x}   積分" & vbCrLf & _
        "     x^2 + y^2 = r^2        二乗" & vbCrLf & _
        "     \begin{pmatrix}a&b\\c&d\end{pmatrix}  行列", _
        "LaTeX 数式挿入", "")
    If latex = "" Then Exit Sub

    Dim scaleStr As String
    scaleStr = InputBox( _
        "解像度スケール（数値が大きいほど高画質）:" & vbCrLf & _
        "2 = 標準  /  4 = 高品質（推奨）  /  6 = 超高精細", _
        "スケール設定", DEFAULT_SCALE)
    If scaleStr = "" Then Exit Sub
    If Not IsNumeric(scaleStr) Then scaleStr = DEFAULT_SCALE

    Call RenderAndInsert(latex, CSng(scaleStr), True)
End Sub

' ---- 挿入済み数式の再編集 ----
Public Sub EditLatex()
    Dim shp As Shape
    On Error Resume Next
    Set shp = Application.ActiveWindow.Selection.ShapeRange(1)
    On Error GoTo 0

    If shp Is Nothing Then
        MsgBox "スライド上の数式画像を選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Dim altText As String
    altText = shp.AlternativeText
    If Left(altText, 6) <> "latex:" Then
        MsgBox "選択したオブジェクトは LaTeX 数式ではありません。", vbExclamation
        Exit Sub
    End If

    Dim oldLatex As String
    oldLatex = Mid(altText, 7)

    Dim newLatex As String
    newLatex = InputBox("LaTeX を編集してください:", "LaTeX 再編集", oldLatex)
    If newLatex = "" Or newLatex = oldLatex Then Exit Sub

    ' 元の位置・サイズを保存して削除
    Dim origLeft As Single : origLeft = shp.Left
    Dim origTop  As Single : origTop  = shp.Top

    Dim parentSlide As Slide
    Set parentSlide = shp.Parent
    shp.Delete

    Call RenderAndInsert(newLatex, CSng(DEFAULT_SCALE), True, origLeft, origTop)
End Sub

' ---- レンダリングして挿入 ----
Private Sub RenderAndInsert(latex As String, scale As Single, display As Boolean, _
    Optional posLeft As Single = -1, Optional posTop As Single = -1)

    Dim exePath As String
    exePath = FindRendererExe()
    If exePath = "" Then Exit Sub

    ' 一時 PNG ファイルパス
    Dim tmpPng As String
    tmpPng = Environ("TEMP") & "\latex_pptx_" & Format(Now, "HHmmss") & ".png"

    ' LaTeX 内のダブルクォートをエスケープ
    Dim safeLatex As String
    safeLatex = Replace(latex, """", """""")

    Dim displayFlag As String
    displayFlag = IIf(display, "1", "0")

    ' コマンド: render.exe "<latex>" <output> <scale> <display>
    Dim cmd As String
    cmd = """" & exePath & """ """ & safeLatex & """ """ & tmpPng & """ " & _
          CStr(scale) & " " & displayFlag

    ' WScript.Shell で同期実行（戻り値を取得）
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim retCode As Long
    retCode = wsh.Run(cmd, 0, True)   ' 0 = ウィンドウ非表示, True = 完了まで待機

    If retCode <> 0 Then
        MsgBox "レンダリングに失敗しました（終了コード: " & retCode & "）" & vbCrLf & _
               "LaTeX 構文を確認してください。" & vbCrLf & vbCrLf & _
               "入力された数式:" & vbCrLf & latex, _
               vbExclamation, "LaTeX エラー"
        Exit Sub
    End If

    If Dir(tmpPng) = "" Then
        MsgBox "PNG ファイルが生成されませんでした。", vbExclamation
        Exit Sub
    End If

    ' 現在のスライドに挿入
    Dim currentSlide As Slide
    On Error Resume Next
    Set currentSlide = Application.ActiveWindow.View.Slide
    On Error GoTo 0

    If currentSlide Is Nothing Then
        MsgBox "スライドが選択されていません。", vbExclamation
        Kill tmpPng
        Exit Sub
    End If

    Dim prs As Presentation
    Set prs = Application.ActivePresentation

    ' デフォルト挿入位置（スライド中央付近）
    Dim insertLeft As Single, insertTop As Single
    If posLeft < 0 Then
        insertLeft = (prs.PageSetup.SlideWidth - 180) / 2
    Else
        insertLeft = posLeft
    End If
    If posTop < 0 Then
        insertTop = (prs.PageSetup.SlideHeight - 80) / 2
    Else
        insertTop = posTop
    End If

    Dim shp As Shape
    Set shp = currentSlide.Shapes.AddPicture( _
        Filename:=tmpPng, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=insertLeft, _
        Top:=insertTop)

    ' LaTeX ソースを alt-text に保存（再編集用）
    shp.AlternativeText = "latex:" & latex

    shp.Select

    ' 一時ファイル削除
    On Error Resume Next
    Kill tmpPng
    On Error GoTo 0
End Sub

' ---- renderer exe を探す ----
Private Function FindRendererExe() As String
    Dim candidates(4) As String
    Dim i As Integer

    ' 1. PPAM と同じフォルダ
    Dim addin As AddIn
    For Each addin In Application.AddIns
        If addin.Loaded Then
            If InStr(LCase(addin.Name), "latexinserter") > 0 Or _
               InStr(LCase(addin.Name), "latex-inserter") > 0 Or _
               InStr(LCase(addin.Name), "latex_inserter") > 0 Then
                candidates(0) = Left(addin.FullName, InStrRev(addin.FullName, "\")) & RENDERER_EX
            End If
        End If
    Next addin

    ' 2. %APPDATA%\Microsoft\AddIns\
    candidates(1) = Environ("APPDATA") & "\Microsoft\AddIns\" & RENDERER_EX

    ' 3. %LOCALAPPDATA%\LaTeXInserter\
    candidates(2) = Environ("LOCALAPPDATA") & "\LaTeXInserter\" & RENDERER_EX

    ' 4. デスクトップ
    candidates(3) = Environ("USERPROFILE") & "\Desktop\" & RENDERER_EX

    For i = 0 To 3
        If candidates(i) <> "" And Dir(candidates(i)) <> "" Then
            FindRendererExe = candidates(i)
            Exit Function
        End If
    Next i

    ' 見つからない場合
    MsgBox RENDERER_EX & " が見つかりません。" & vbCrLf & vbCrLf & _
           "以下のいずれかに配置してください:" & vbCrLf & _
           "  ・PPAM ファイルと同じフォルダ" & vbCrLf & _
           "  ・%APPDATA%\Microsoft\AddIns\" & vbCrLf & vbCrLf & _
           "GitHub Releases からダウンロード:" & vbCrLf & _
           "https://github.com/yuuukichiiii/latex-eq-pptx/releases", _
           vbExclamation, "レンダラーが見つかりません"
    FindRendererExe = ""
End Function
