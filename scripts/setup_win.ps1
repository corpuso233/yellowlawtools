$GREEN = "Green"; $YELLOW = "Yellow"; $RED = "Red"
$DESKTOP = [Environment]::GetFolderPath("Desktop")
$TOOLS_DIR = "$DESKTOP\YellowLawTools"
$REPO = "https://raw.githubusercontent.com/corpuso233/yellowlawtools/main/tools"

Clear-Host
Write-Host "==============================================" -ForegroundColor Yellow
Write-Host "   Yellow Law Group PC - Arac Kurulumu" -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Hangi araclari kurmak istiyorsunuz?"
Write-Host "  1) EOIR A-Number Sorgulama"
Write-Host "  2) USCIS Case Status Sorgulama"
Write-Host "  3) Annual Asylum Fee Kontrolu"
Write-Host "  4) Hepsini Kur"
Write-Host ""
$SECIM = Read-Host "Seciminiz"

# Python kontrolü
Write-Host "Python kontrol ediliyor..." -ForegroundColor Yellow
$pythonOK = $false
try { python --version 2>&1 | Out-Null; $pythonOK = $true } catch {}
try { python3 --version 2>&1 | Out-Null; $pythonOK = $true } catch {}

if (-not $pythonOK) {
    Write-Host "Python bulunamadi! python.org/downloads adresinden yukleyin." -ForegroundColor Red
    Read-Host "ENTER'a basin"
    exit
}
Write-Host "OK Python mevcut" -ForegroundColor Green

# Paketler
Write-Host "Paketler yukleniyor..." -ForegroundColor Yellow
pip install selenium openpyxl webdriver-manager playwright -q 2>$null
python -m playwright install chromium 2>$null
Write-Host "OK Paketler yuklendi" -ForegroundColor Green

New-Item -ItemType Directory -Force -Path $TOOLS_DIR | Out-Null

function Install-Tool($file, $display, $headers) {
    Write-Host "$display kuruluyor..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$REPO/$file" -OutFile "$TOOLS_DIR\$file" -ErrorAction SilentlyContinue

    if (-not (Test-Path "$TOOLS_DIR\$file") -or (Get-Item "$TOOLS_DIR\$file").Length -eq 0) {
        Write-Host "Indirme basarisiz!" -ForegroundColor Red
        return
    }

    $bat = "$TOOLS_DIR\$display.bat"
    "@echo off`ncd `"$TOOLS_DIR`"`npython `"$file`"`npause" | Set-Content $bat
    Write-Host "OK $display kuruldu" -ForegroundColor Green
}

function Make-Excel($filename, $headers) {
    $headerList = $headers -split "\|"
    python -c @"
import openpyxl, os
wb = openpyxl.Workbook()
ws = wb.active
headers = '$headers'.split('|')
for i, h in enumerate(headers, 1):
    ws.cell(row=1, column=i).value = h.strip()
wb.save(r'$DESKTOP\$filename')
print('Excel sablonu olusturuldu: $filename')
"@
}

if ($SECIM -match "1" -or $SECIM -eq "4") {
    Install-Tool "eoir_otomasyon.py" "EOIR_A-Number_Sorgulama" "A-Number|Sonuc"
    Make-Excel "eoir_anumbers.xlsx" "A-Number|Sonuc"
}
if ($SECIM -match "2" -or $SECIM -eq "4") {
    Install-Tool "uscis_case_status.py" "USCIS_Case_Status" "Receipt Number|Case Status|Annual Fee Notu|Aciklama"
    Make-Excel "uscis_receipt_numbers.xlsx" "Receipt Number|Case Status|Annual Fee Notu|Aciklama"
}
if ($SECIM -match "3" -or $SECIM -eq "4") {
    Install-Tool "annual_fee_check.py" "Annual_Fee_Kontrolu" "A-Number|Receipt Number|Sonuc"
    Make-Excel "annual_fee.xlsx" "A-Number|Receipt Number|Sonuc"
}

Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host "OK Kurulum tamamlandi!" -ForegroundColor Green
Write-Host ""
Write-Host "Masaustundeki 'YellowLawTools' klasorune gidin."
Write-Host "Ilgili .bat dosyasina cift tiklayin."
Write-Host "Excel sablonlari masaustune kaydedildi."
Write-Host "==============================================" -ForegroundColor Green
Read-Host "Cikmak icin ENTER'a basin"
