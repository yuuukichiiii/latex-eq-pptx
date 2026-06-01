# LaTeX Equation Inserter for PowerPoint

> **⚠️ 開発中 (Work in Progress)**
> このプロジェクトは現在開発中です。動作はしますが、インストール手順が煩雑な部分があります。
> PPAM ファイルの自動生成・配布は未対応のため、手動セットアップが必要です。

PowerPoint スライドに LaTeX 数式を PNG 画像として挿入するアドイン。
**TeX Live・インターネット接続不要**、企業の Microsoft 365 環境でも動作。

---

## 動作環境

- Windows 10 / 11
- Microsoft 365 または PowerPoint 2016 以降
- インターネット接続：**不要**（完全オフライン動作）

---

## セットアップ手順（PPAM 方式 ― 推奨）

### 1. latex-renderer.exe をダウンロード

[GitHub Releases](https://github.com/yuuukichiiii/latex-eq-pptx/releases/latest) から
`latex-renderer.exe`（約 40 MB）をダウンロードし、以下に配置：

```
C:\Users\<ユーザー名>\AppData\Roaming\Microsoft\AddIns\latex-renderer.exe
```

> **補足**: `%APPDATA%\Microsoft\AddIns\` をエクスプローラのアドレスバーに貼り付けると開けます。

---

### 2. PowerPoint で VBA モジュールを追加

#### 2-1. 開発タブを表示

**ファイル → オプション → リボンのユーザー設定** →
右側の「開発」にチェックを入れる → OK

#### 2-2. VBA エディタを開く

PowerPoint で適当なファイルを開き、**開発タブ → Visual Basic**（または `Alt + F11`）

#### 2-3. モジュールを追加

メニュー **挿入 → 標準モジュール** → 空のコードウィンドウが開く

#### 2-4. コードを貼り付け

以下のコードを全選択してコピーし、コードウィンドウに貼り付ける：

```vba
Option Explicit

Public Sub InsertLatex()
    Dim latex As String
    Dim sc As String
    Dim tmp As String
    Dim exe As String
    Dim cmd As String
    Dim ret As Long
    Dim sl As Slide
    Dim shp As Shape

    latex = InputBox("LaTeX を入力してください", "LaTeX 数式挿入", "\frac{a}{b}")
    If latex = "" Then Exit Sub

    sc = InputBox("スケール（推奨: 4）", "解像度", "4")
    If sc = "" Then sc = "4"

    exe = Environ("APPDATA") & "\Microsoft\AddIns\latex-renderer.exe"
    If Dir(exe) = "" Then
        exe = Environ("USERPROFILE") & "\Desktop\latex-renderer.exe"
    End If
    If Dir(exe) = "" Then
        MsgBox "latex-renderer.exe が見つかりません" & vbCrLf & _
               "場所: " & Environ("APPDATA") & "\Microsoft\AddIns\"
        Exit Sub
    End If

    tmp = Environ("TEMP") & "\latex_out.png"
    cmd = """" & exe & """ """ & latex & """ """ & tmp & """ " & sc & " 1"

    ret = CreateObject("WScript.Shell").Run(cmd, 0, True)
    If ret <> 0 Then
        MsgBox "レンダリング失敗 (コード: " & ret & ")" & vbCrLf & latex
        Exit Sub
    End If

    Set sl = Application.ActiveWindow.View.Slide
    Set shp = sl.Shapes.AddPicture(tmp, False, True, 100, 100)
    shp.AlternativeText = "latex:" & latex
    shp.Select
End Sub
```

---

### 3. PPAM として保存（マクロを永続化）

VBA エディタで **ファイル → 名前を付けて保存**

| 項目 | 設定 |
|---|---|
| 保存先 | `C:\Users\<ユーザー名>\AppData\Roaming\Microsoft\AddIns\` |
| ファイル名 | `LaTeXInserter` |
| ファイルの種類 | **PowerPoint アドイン (*.ppam)** |

→ **保存**

---

### 4. PPAM を PowerPoint に登録

PowerPoint を**再起動**してから：

**ファイル → オプション → アドイン**
→ 管理: **PowerPoint アドイン** → **設定...**
→ **追加** → `LaTeXInserter.ppam` を選択 → OK

---

### 5. ボタンを追加（省略可）

毎回 `Alt + F8` を押す手間を省くため、クイックアクセスツールバーに登録：

PowerPoint 左上の **▼** → **その他のコマンド**
→ コマンドの選択: **マクロ**
→ `InsertLatex` を選択 → **追加** → OK

---

## 使い方

1. PowerPoint でスライドを開く
2. ツールバーの **InsertLatex ボタン**（または `Alt + F8` → InsertLatex → 実行）
3. LaTeX 数式を入力して OK

   ```
   例: \frac{a^2 + b^2}{c^2}
       \int_0^\infty e^{-x}\,dx
       \begin{pmatrix}a & b \\ c & d\end{pmatrix}
   ```

4. スケール（解像度）を入力して OK（推奨: 4）
5. 数式がスライドに挿入される

---

## LaTeX 記法の例

| 表示 | LaTeX |
|---|---|
| 分数 | `\frac{a}{b}` |
| 積分 | `\int_0^\infty f(x)\,dx` |
| 総和 | `\sum_{n=1}^{\infty} a_n` |
| 行列 | `\begin{pmatrix}a & b \\ c & d\end{pmatrix}` |
| 平方根 | `\sqrt{x^2 + y^2}` |
| 極限 | `\lim_{x \to 0} \frac{\sin x}{x}` |
| オイラー | `e^{i\pi} + 1 = 0` |

---

## 技術仕様

### レンダリングパイプライン

```
LaTeX 文字列
  └─→ MathJax 3 (liteAdaptor, ブラウザ不要)
        └─→ SVG (フォントをパスに変換、自己完結)
              └─→ @resvg/resvg-js (Rust 製 SVG レンダラ)
                    └─→ PNG (4× スケール ≈ 300 DPI)
                          └─→ PowerPoint スライドに挿入
```

- **MathJax 3**: TeX Live 不要、Node.js で LaTeX → SVG 変換
- **resvg-js**: Rust 製の高速 SVG → PNG 変換
- **pkg**: Node.js + 依存関係を単一 exe にバンドル

### ファイル構成

```
latex-eq-pptx/
├── renderer/          # latex-renderer.exe のソース
│   ├── render.js      # MathJax + resvg-js によるレンダリング
│   └── package.json
├── vba/
│   └── LatexInserter.bas   # PowerPoint VBA コード
├── src/taskpane/      # Web アドイン版（開発中）
├── manifest.xml       # Web アドイン用マニフェスト
└── .github/workflows/
    ├── deploy.yml     # GitHub Pages へのデプロイ
    └── build-exe.yml  # latex-renderer.exe のビルド
```

---

## 既知の問題 / 未実装

- [ ] PPAM ファイルの自動生成・配布（現状は手動セットアップが必要）
- [ ] 挿入済み数式の再編集機能（EditLatex）
- [ ] リボンへの自動ボタン登録
- [ ] Web アドイン版（企業 IT ポリシーで利用不可の環境があるため PPAM 版を推奨）
- [ ] Mac 対応

---

## 開発者向け

### renderer のビルド

```bash
cd renderer
npm install
npm run build   # → dist-exe/latex-renderer.exe
```

> **注意**: MathJax 3.2.2 に Node.js v20+ との互換性バグあり。
> `renderer/node_modules/mathjax-full/js/input/tex/base/BaseConfiguration.js` の
> `range[4] &&` を `range && range[4] &&` に修正が必要（GitHub Actions で自動適用済み）。

### GitHub Actions

- **push to master** → GitHub Pages に Web アドイン版をデプロイ
- **push tag `v*`** → `latex-renderer.exe` をビルドして GitHub Release に添付

---

## ライセンス

MIT
