# Install-LaTeXAddin.ps1
# PowerPoint LaTeX 数式アドインのインストールスクリプト
# 管理者権限不要・インターネット接続が必要

param(
    [switch]$Uninstall
)

$AddinName   = "LaTeXPPTXAddin"
$ManifestUrl = "https://yuuukichiiii.github.io/latex-eq-pptx/manifest.xml"
$AddinDir    = "$env:LOCALAPPDATA\$AddinName"
$ShareName   = $AddinName
$RegBase     = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$CatalogId   = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
$RegPath     = "$RegBase\$CatalogId"

# ---- アンインストール ----
if ($Uninstall) {
    Write-Host "アドインを削除します..."
    if (Test-Path $RegPath) { Remove-Item -Path $RegPath -Force }
    try { net share "$ShareName" /delete 2>$null | Out-Null } catch {}
    if (Test-Path $AddinDir) { Remove-Item -Path $AddinDir -Recurse -Force }
    Write-Host "削除完了。PowerPoint を再起動してください。"
    exit
}

Write-Host ""
Write-Host "===== PowerPoint LaTeX 数式アドイン インストーラー =====" -ForegroundColor Cyan
Write-Host ""

# ---- Step 1: manifest.xml をダウンロード ----
Write-Host "[1/3] manifest.xml をダウンロード中..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path $AddinDir | Out-Null
try {
    Invoke-WebRequest -Uri $ManifestUrl -OutFile "$AddinDir\manifest.xml" -UseBasicParsing
    Write-Host "      OK: $AddinDir\manifest.xml" -ForegroundColor Green
} catch {
    Write-Host "エラー: manifest.xml のダウンロードに失敗しました。" -ForegroundColor Red
    Write-Host "インターネット接続を確認してください: $ManifestUrl"
    exit 1
}

# ---- Step 2: フォルダを共有 ----
Write-Host "[2/3] フォルダを共有中..." -ForegroundColor Yellow
$null = net share "$ShareName=$AddinDir" /grant:$env:USERNAME,full 2>&1
$UncPath = "\\$env:COMPUTERNAME\$ShareName"
Write-Host "      OK: $UncPath" -ForegroundColor Green

# ---- Step 3: レジストリに信頼済みカタログとして登録 ----
Write-Host "[3/3] PowerPoint に信頼済みカタログとして登録中..." -ForegroundColor Yellow
if (-not (Test-Path $RegBase)) { New-Item -Path $RegBase -Force | Out-Null }
New-Item -Path $RegPath -Force | Out-Null
Set-ItemProperty -Path $RegPath -Name "Id"    -Value $CatalogId
Set-ItemProperty -Path $RegPath -Name "Url"   -Value $UncPath
Set-ItemProperty -Path $RegPath -Name "Flags" -Value 1 -Type DWord
Write-Host "      OK: レジストリ登録完了" -ForegroundColor Green

# ---- 完了メッセージ ----
Write-Host ""
Write-Host "===== インストール完了 =====" -ForegroundColor Cyan
Write-Host ""
Write-Host "次の手順でアドインを有効化してください:" -ForegroundColor White
Write-Host ""
Write-Host "  1. PowerPoint を完全に終了して再起動"
Write-Host "  2. [挿入] タブ → [マイ アドイン] をクリック"
Write-Host "     ※ ボタンが無い場合: [挿入] → 右端の [アドイン] 小さいボタン"
Write-Host "  3. [共有フォルダー] タブ → [LaTeX 数式挿入] → [追加]"
Write-Host ""
Write-Host "アンインストールするには: .\Install-LaTeXAddin.ps1 -Uninstall" -ForegroundColor Gray
Write-Host ""
