# ============================================================
# USCIS Case Status Otomasyonu
# Yellow Law Group PC
# ============================================================
import os, time, subprocess, openpyxl
from playwright.sync_api import sync_playwright

DESKTOP        = os.path.join(os.path.expanduser("~"), "Desktop")
EXCEL_DOSYASI  = os.path.join(DESKTOP, "uscis_receipt_numbers.xlsx")
USCIS_URL      = "https://egov.uscis.gov/"

# Chrome yolu (Mac ve Windows otomatik tespit)
if os.name == "nt":  # Windows
    CHROME_YOLU = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
else:  # Mac
    CHROME_YOLU = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

RECEIPT_SUTUN    = 1
SONUC_SUTUN      = 2
ANNUAL_FEE_SUTUN = 3
ACIKLAMA_SUTUN   = 4
BASLIK_SATIRI    = 1
MAX_DENEME       = 3
BEKLEME_SANIYE   = 12

ANNUAL_FEE_KELIMELER = [
    "annual asylum fee", "annual fee", "asylum fee",
    "fee payment", "pay your", "pay the",
]

NAVIGASYON_METINLER = {
    "topics", "forms", "newsroom", "citizenship", "green card",
    "laws", "tools", "case status online", "check case status",
    "uscis", "check status", "popular topics", "related links",
    "related tools", "enter a receipt number", "enter another receipt number",
    "already have an account", "dhs privacy notice", "paperwork reduction act"
}


def receipt_gir_ve_gonder(page, receipt_number):
    for sec in [
        "input[placeholder='EAC1234567890']",
        "input[placeholder*='1234']",
        "input[id*='receipt']",
        "input[name*='receipt']",
        "input[type='text']",
    ]:
        try:
            el = page.locator(sec).first
            if el.is_visible(timeout=3000):
                el.click()
                page.wait_for_timeout(200)
                el.fill("")
                page.wait_for_timeout(200)
                el.fill(receipt_number)
                page.wait_for_timeout(1500)
                for btn_sec in ["button:has-text('Check Status')", "button[type='submit']", "input[type='submit']"]:
                    try:
                        btn = page.locator(btn_sec).first
                        btn.wait_for(state="visible", timeout=4000)
                        btn.click()
                        return True
                    except Exception:
                        continue
        except Exception:
            continue
    return False


def sonuc_var_mi(page):
    try:
        for el in page.locator("h2").all():
            try:
                metin = el.inner_text().strip()
                metin_lower = metin.lower()
                if (metin and len(metin) > 8
                        and metin_lower not in NAVIGASYON_METINLER
                        and not any(x in metin_lower for x in [
                            "check", "enter", "receipt", "case status online",
                            "topic", "form", "news", "citizenship", "green card",
                            "law", "tool", "login", "privacy"])):
                    return True
            except Exception:
                continue
    except Exception:
        pass
    for hata_sec in ["text=invalid", "text=please try again", "text=not found"]:
        try:
            if page.locator(hata_sec).first.is_visible(timeout=300):
                return True
        except Exception:
            pass
    return False


def sonucu_oku(page):
    baslik = None
    for el in page.locator("h2").all():
        try:
            metin = el.inner_text().strip()
            metin_lower = metin.lower()
            if (metin and len(metin) > 8
                    and metin_lower not in NAVIGASYON_METINLER
                    and not any(x in metin_lower for x in [
                        "check", "enter", "receipt", "case status online",
                        "topic", "form", "news", "citizenship", "green card",
                        "law", "tool", "login", "privacy"])):
                baslik = metin
                break
        except Exception:
            continue

    if not baslik:
        for hata_sec in ["text=invalid", "text=please try again", "text=not found"]:
            try:
                if page.locator(hata_sec).first.is_visible(timeout=300):
                    return "HATA: Receipt geçersiz veya bulunamadı", "", False
            except Exception:
                pass
        return "Sonuç alınamadı - manuel kontrol et", "", False

    try:
        aciklama = page.locator("div.caseStatusSection div.conditionalLanding p").first.inner_text(timeout=3000).strip()
    except Exception:
        try:
            aciklama = page.locator("#landing-page-header + p, #landing-page-header ~ p").first.inner_text(timeout=2000).strip()
        except Exception:
            aciklama = ""

    sayfa_metni = page.inner_text("body")
    annual_fee_var = any(k in sayfa_metni.lower() for k in ANNUAL_FEE_KELIMELER)
    return baslik, aciklama, annual_fee_var


def receipt_sorgula(page, receipt_number):
    for deneme in range(1, MAX_DENEME + 1):
        if deneme > 1:
            print(f"   → ↺ {deneme}. deneme...")
            page.goto(USCIS_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

        if not receipt_gir_ve_gonder(page, receipt_number):
            continue

        for _ in range(BEKLEME_SANIYE):
            page.wait_for_timeout(1000)
            if sonuc_var_mi(page):
                return sonucu_oku(page)

        print(f"   → ⏱ Cevap gelmedi...")

    return "Sonuç alınamadı - manuel kontrol et", "", False


def main():
    print("=" * 55)
    print("  USCIS Case Status Otomasyonu")
    print("  Yellow Law Group PC")
    print("=" * 55)
    print(f"\nExcel aranıyor: {EXCEL_DOSYASI}")

    try:
        wb = openpyxl.load_workbook(EXCEL_DOSYASI)
        ws = wb.active
    except FileNotFoundError:
        print(f"\nHATA: 'uscis_receipt_numbers.xlsx' masaüstünde bulunamadı!")
        input("\nÇıkmak için ENTER'a basın...")
        return

    ws.cell(row=1, column=SONUC_SUTUN).value      = "Case Status"
    ws.cell(row=1, column=ANNUAL_FEE_SUTUN).value = "Annual Fee Notu"
    ws.cell(row=1, column=ACIKLAMA_SUTUN).value   = "Açıklama"

    toplam = sum(1 for row in ws.iter_rows(min_row=2) if row[RECEIPT_SUTUN - 1].value)
    print(f"Toplam {toplam} Receipt Number bulundu.\n")

    # Chrome'u kapat ve debug modunda aç
    if os.name == "nt":
        subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], capture_output=True)
    else:
        subprocess.run(["pkill", "-9", "-f", "Google Chrome"], capture_output=True)
    time.sleep(3)

    subprocess.Popen([CHROME_YOLU,
        "--remote-debugging-port=9222",
        "--no-first-run",
        "--no-default-browser-check",
        "--user-data-dir=" + os.path.join(os.path.expanduser("~"), ".uscis_debug_profil"),
    ])
    time.sleep(4)

    print("\n>>> Chrome açıldı. egov.uscis.gov'a git, Cloudflare çıkarsa geç.")
    input("Hazır olunca ENTER'a bas...")

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            print("✅ Chrome'a bağlandı!\n")
        except Exception as e:
            print(f"❌ Bağlanamadı: {e}")
            input("ENTER'a basın...")
            return

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(USCIS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)

        islenen = 0
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            receipt_h = row[RECEIPT_SUTUN - 1]
            sonuc_h   = row[SONUC_SUTUN - 1]
            fee_h     = row[ANNUAL_FEE_SUTUN - 1]
            acik_h    = row[ACIKLAMA_SUTUN - 1]
            if not receipt_h.value:
                continue

            receipt = str(receipt_h.value).strip().upper()
            islenen += 1
            print(f"[{islenen}/{toplam}] {receipt}")

            try:
                sonuc, aciklama, fee = receipt_sorgula(page, receipt)
                sonuc_h.value = sonuc
                acik_h.value  = aciklama
                fee_h.value   = "⚠️ ANNUAL FEE GEREKLİ" if fee else ""
                print(f"   → ✅ {sonuc}")
                if fee:
                    print(f"   → 💰 Annual Fee tespit edildi!")
            except Exception as e:
                sonuc_h.value = f"HATA: {str(e)[:80]}"
                print(f"   → ❌ {e}")

            wb.save(EXCEL_DOSYASI)

    print(f"\n✅ Tamamlandı! {islenen} kayıt işlendi.")
    print(f"Excel: {EXCEL_DOSYASI}")
    input("\nÇıkmak için ENTER'a basın...")


if __name__ == "__main__":
    main()
