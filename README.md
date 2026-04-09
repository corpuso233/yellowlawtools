# Yellow Law Group PC - Tools

Immigration law practice automation tools.

## Kurulum / Installation

### Mac
Terminal'i açın ve şunu çalıştırın:
```bash
curl -fsSL https://raw.githubusercontent.com/corpuso233/yellowlawtools/main/scripts/setup_mac.sh | bash
```

### Windows
PowerShell'i **Yönetici olarak** açın ve şunu çalıştırın:
```powershell
irm https://raw.githubusercontent.com/corpuso233/yellowlawtools/main/scripts/setup_win.ps1 | iex
```

---

## Araçlar / Tools

### 1. EOIR A-Number Sorgulama
- Excel'e A-Number'ları girin (A sütunu, tiresiz)
- Script EOIR portala giriş yapar
- Sonuçları B sütununa yazar

### 2. USCIS Case Status Sorgulama  
- Excel'e Receipt Number'ları girin (A sütunu)
- Script USCIS'te sorgular
- Case Status → B sütunu
- Annual Fee notu → C sütunu
- Açıklama → D sütunu

### 3. Annual Asylum Fee Kontrolü
- Excel'e A-Number (A) ve Receipt Number (B) girin
- Script my.uscis.gov'da kontrol eder
- Sonuç → C sütunu

---

## Notlar
- Excel dosyaları masaüstünde olmalıdır
- EOIR için portal girişi gereklidir (2FA dahil)
- USCIS için giriş gerekmez
- Annual Fee için my.uscis.gov girişi gereklidir
