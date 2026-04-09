# ============================================================
# Annual Asylum Fee Kontrol Otomasyonu
# Yellow Law Group PC
# ============================================================
import os, time, subprocess, openpyxl
from playwright.sync_api import sync_playwright

DESKTOP           = os.path.join(os.path.expanduser("~"), "Desktop")
EXCEL_DOSYASI     = os.path.join(DESKTOP, "annual_fee.xlsx")
QUESTIONNAIRE_URL = "https://my.uscis.gov/accounts/annual-asylum-fee/questionnaire"

if os.name == "nt":
    CHROME_YOLU = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
else:
    CHROME_YOLU = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

A_NUMBER_SUTUN = 1
RECEIPT_SUTUN  = 2
SONUC_SUTUN    = 3
BASLIK_SATIRI  = 1
MAX_DENEME     = 3
BEKLEME_SANIYE = 12


def next_varsa_bas(page):
    try:
        btn = page.locator("button:has-text('Next')").last
        if btn.is_visible(timeout=2000):
            btn.scroll_into_view_if_needed()
            page.wait_for_timeout(500)
            btn.click()
            page.wait_for_timeout(2000)
            print("   → Next'e basıldı ✓")
    except Exception:
        pass


def form_doldur_ve_gonder(page, a_number, receipt_number):
    receipt_str = str(receipt_number).strip()
    if len(receipt_str) != 13:
        raise ValueError(f"Receipt Number geçersiz uzunluk: '{receipt_str}' ({len(receipt_str)} karakter, 13 olmalı)")

    a_input = None
    for sec in ["input[id*='alien']", "input[id*='aNumber']", "input[placeholder*='A-']", "input[aria-label*='A-Number']"]:
        try:
            el = page.locator(sec).first
            if el.is_visible(timeout=2000):
                a_input = el
                break
        except Exception:
            continue

    if not a_input:
        try:
            inputs = page.locator("input[type='text'], input:not([type])").all()
            if inputs:
                a_input = inputs[0]
        except Exception:
            pass

    if not a_input:
        return False, "A-Number input bulunamadı"

    a_input.click()
    page.wait_for_timeout(200)
    a_input.fill("")
    page.wait_for_timeout(200)
    a_input.fill(str(a_number).strip())
    page.wait_for_timeout(500)

    receipt_input = None
    for sec in ["input[placeholder='EAC1234567890']", "input[placeholder*='receipt' i]", "input[aria-label*='receipt' i]", "input[id*='receipt']"]:
        try:
            el = page.locator(sec).first
            if el.is_visible(timeout=2000):
                receipt_input = el
                break
        except Exception:
            continue

    if not receipt_input:
        try:
            inputs = page.locator("input[type='text'], input:not([type])").all()
            if len(inputs) >= 2:
                receipt_input = inputs[1]
        except Exception:
            pass

    if not receipt_input:
        return False, "Receipt input bulunamadı"

    receipt_input.click()
    page.wait_for_timeout(200)
    receipt_input.fill("")
    page.wait_for_timeout(200)
    receipt_input.fill(receipt_str)
    page.wait_for_timeout(1000)

    for btn_sec in ["button:has-text('Continue to payment')", "button:has-text('Continue')", "button[type='submit']", "input[type='submit']"]:
        try:
            btn = page.locator(btn_sec).last
            btn.wait_for(state="visible", timeout=4000)
            btn.scroll_into_view_if_needed()
            page.wait_for_timeout(500)
            btn.click()
            return True, ""
        except Exception:
            continue

    return False, "Continue butonu bulunamadı"


def sayfa_hazir_mi(page):
    try:
        metin = page.inner_text("body").lower()
        return any(x in metin for x in [
            "annual asylum fee payment is not due",
            "pay for your annual asylum fee",
            "could not find your case",
            "payment is due"
        ])
    except Exception:
        return False


def sonucu_oku(page):
    try:
        metin = page.inner_text("body")
        metin_lower = metin.lower()

        if "pay for your annual asylum fee" in metin_lower:
            try:
                parcalar = []
                for p in page.locator("div[style*='padding-bottom'] p").all():
                    try:
                        t = p.inner_text(timeout=1000).strip()
                        if t:
                            parcalar.append(t)
                    except Exception:
                        continue
                if parcalar:
                    return f"⚠️ ÖDEME GEREKLİ: {' | '.join(parcalar[:3])}"
            except Exception:
                pass
            return "⚠️ ÖDEME GEREKLİ: Pay For Your Annual Asylum Fee sayfası açıldı"

        if "could not find your case" in metin_lower:
            return "HATA: We could not find your case - A-Number veya Receipt Number hatalı"

        if "annual asylum fee payment is not due" in metin_lower:
            try:
                uyari = page.locator("#case-paid-for-alert p").first.inner_text(timeout=3000).strip()
                if uyari:
                    return f"ÖDEME GEREKMİYOR: {uyari}"
            except Exception:
                pass
            return "ÖDEME GEREKMİYOR: Annual Asylum Fee henüz ödenmesi gerekmiyor"

        if "payment is due" in metin_lower:
            return "⚠️ ÖDEME GEREKLİ"
    except Exception:
        pass
    return None


def bir_kayit_isle(page, a_number, receipt_number):
    receipt_str_kontrol = str(receipt_number).strip()
    if len(receipt_str_kontrol) != 13:
        return f"Receipt Number geçersiz uzunluk: '{receipt_str_kontrol}' ({len(receipt_str_kontrol)} karakter, 13 olmalı)"

    for deneme in range(1, MAX_DENEME + 1):
        if deneme > 1:
            print(f"   → ↺ {deneme}. deneme...")
        try:
            page.goto(QUESTIONNAIRE_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)
            next_varsa_bas(page)

            basarili, hata = form_doldur_ve_gonder(page, a_number, receipt_number)
            if not basarili:
                print(f"   → ⚠️ {hata}")
                continue

            for _ in range(BEKLEME_SANIYE):
                page.wait_for_timeout(1000)
                sonuc = sonucu_oku(page)
                if sonuc:
                    return sonuc

            print(f"   → ⏱ Cevap gelmedi...")
        except ValueError as e:
            return str(e)
        except Exception as e:
            print(f"   → ❌ {e}")
            continue

    return "Sonuç alınamadı - manuel kontrol et"


def main():
    print("=" * 55)
    print("  Annual Asylum Fee Kontrol Otomasyonu")
    print("  Yellow Law Group PC")
    print("=" * 55)
    print(f"\nExcel aranıyor: {EXCEL_DOSYASI}")

    try:
        wb = openpyxl.load_workbook(EXCEL_DOSYASI)
        ws = wb.active
    except FileNotFoundError:
        print(f"\nHATA: 'annual_fee.xlsx' masaüstünde bulunamadı!")
        input("\nÇıkmak için ENTER'a basın...")
        return

    ws.cell(row=1, column=A_NUMBER_SUTUN).value = "A-Number"
    ws.cell(row=1, column=RECEIPT_SUTUN).value  = "Receipt Number"
    ws.cell(row=1, column=SONUC_SUTUN).value    = "Sonuç"

    toplam = sum(1 for row in ws.iter_rows(min_row=2)
                 if row[A_NUMBER_SUTUN - 1].value and row[RECEIPT_SUTUN - 1].value)
    print(f"Toplam {toplam} kayıt bulundu.\n")

    if os.name == "nt":
        subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], capture_output=True)
    else:
        subprocess.run(["pkill", "-9", "-f", "Google Chrome"], capture_output=True)
    time.sleep(3)

    subprocess.Popen([CHROME_YOLU,
        "--remote-debugging-port=9222",
        "--no-first-run",
        "--no-default-browser-check",
        "--user-data-dir=" + os.path.join(os.path.expanduser("~"), ".annual_fee_profil"),
    ])
    time.sleep(4)

    print("\n>>> Chrome açıldı. my.uscis.gov'a git, giriş yap.")
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

        islenen = 0
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            a_h       = row[A_NUMBER_SUTUN - 1]
            receipt_h = row[RECEIPT_SUTUN - 1]
            sonuc_h   = row[SONUC_SUTUN - 1]
            if not a_h.value or not receipt_h.value:
                continue

            a_number = str(a_h.value).strip()
            receipt  = str(receipt_h.value).strip().upper()
            islenen += 1
            print(f"[{islenen}/{toplam}] A: {a_number} | Receipt: {receipt}")

            try:
                sonuc = bir_kayit_isle(page, a_number, receipt)
                sonuc_h.value = sonuc
                print(f"   → {sonuc}")
            except Exception as e:
                sonuc_h.value = f"HATA: {str(e)[:80]}"
                print(f"   → ❌ {e}")

            wb.save(EXCEL_DOSYASI)

    print(f"\n✅ Tamamlandı! {islenen} kayıt işlendi.")
    print(f"Excel: {EXCEL_DOSYASI}")
    input("\nÇıkmak için ENTER'a basın...")


if __name__ == "__main__":
    main()
