# Create-PPAM.ps1
# LaTeXInserter.ppam を自動生成するスクリプト
# 実行方法: PowerShell で右クリック → PowerShell で実行

$ErrorActionPreference = 'Stop'

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$BasFile    = Join-Path $ScriptDir "vba\LatexInserter.bas"
$PpamOutput = Join-Path $ScriptDir "LaTeXInserter.ppam"

Write-Host ""
Write-Host "===== LaTeXInserter.ppam 自動生成 =====" -ForegroundColor Cyan
Write-Host ""

# ---- .bas ファイルの確認 ----
if (-not (Test-Path $BasFile)) {
    Write-Host "エラー: $BasFile が見つかりません。" -ForegroundColor Red
    Write-Host "このスクリプトはリポジトリのルートフォルダで実行してください。"
    pause; exit 1
}

$vbaCode = Get-Content $BasFile -Raw -Encoding UTF8

# ---- PowerPoint COM で PPAM を作成 ----
Write-Host "[1/4] PowerPoint を起動中..." -ForegroundColor Yellow
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

Write-Host "[2/4] 新しいプレゼンテーションを作成中..." -ForegroundColor Yellow
$prs = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)

Write-Host "[3/4] VBA モジュールを追加中..." -ForegroundColor Yellow
try {
    $vbaProject = $prs.VBProject
    $module = $vbaProject.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
    $module.Name = "LatexInserter"
    $module.CodeModule.DeleteLines(1, $module.CodeModule.CountOfLines)
    $module.CodeModule.AddFromString($vbaCode)
    Write-Host "      OK: VBA モジュール追加完了" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "VBA プロジェクトへのアクセスが拒否されました。" -ForegroundColor Red
    Write-Host ""
    Write-Host "【手動で有効化が必要】" -ForegroundColor Yellow
    Write-Host "PowerPoint → ファイル → オプション → トラスト センター"
    Write-Host "→ トラスト センターの設定 → マクロの設定"
    Write-Host "→ 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック"
    Write-Host ""
    Write-Host "設定後、このスクリプトを再実行してください。"
    $prs.Close()
    $ppt.Quit()
    pause; exit 1
}

Write-Host "[4/4] PPAM として保存中..." -ForegroundColor Yellow
# ppSaveAsOpenXMLAddIn = 30
$prs.SaveAs($PpamOutput, 30)
$prs.Close()
$ppt.Quit()

Write-Host ""
Write-Host "===== 完了 =====" -ForegroundColor Cyan
Write-Host ""
Write-Host "生成されたファイル: $PpamOutput" -ForegroundColor Green
Write-Host ""
Write-Host "次のステップ:"
Write-Host "  1. latex-renderer.exe を同じフォルダ ($ScriptDir) に配置"
Write-Host "  2. PowerPoint を起動"
Write-Host "  3. ファイル → オプション → アドイン"
Write-Host "  4. 管理: PowerPoint アドイン → 設定..."
Write-Host "  5. 追加 → LaTeXInserter.ppam を選択"
Write-Host ""
pause
