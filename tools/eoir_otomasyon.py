# ============================================================
# EOIR A-Number Kontrol Otomasyonu
# Yellow Law Group PC
# ============================================================
import os, time, openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

DESKTOP        = os.path.join(os.path.expanduser("~"), "Desktop")
EXCEL_DOSYASI  = os.path.join(DESKTOP, "eoir_anumbers.xlsx")
PORTAL_URL     = "https://case-access.eoir.justice.gov/addappearanceform/E28"

A_NUMBER_SUTUN      = 1
SONUC_SUTUN         = 2
BASLIK_SATIRI       = 1
BEKLEME_SURESI      = 15
SATIR_ARASI_BEKLEME = 3


def a_numara_formatla(numara):
    numara = str(numara).strip().replace("-", "").replace(" ", "")
    numara = numara.zfill(9)
    if len(numara) == 9:
        return f"{numara[:3]}-{numara[3:6]}-{numara[6:]}"
    return numara


def agree_varsa_bas(driver):
    try:
        agree_btn = driver.find_element(By.XPATH,
            "//button[normalize-space()='Agree'] | //a[normalize-space()='Agree'] | //input[@value='Agree']")
        driver.execute_script("arguments[0].click();", agree_btn)
        print("   → Agree otomatik kabul edildi ✓")
        time.sleep(2)
        return True
    except NoSuchElementException:
        return False


def continue_butonuna_bas(driver):
    yontemler = [
        (By.XPATH, "//button[normalize-space()='Continue']"),
        (By.XPATH, "//button[contains(text(),'Continue')]"),
        (By.XPATH, "//input[@value='Continue']"),
        (By.XPATH, "//button[@type='submit']"),
        (By.CSS_SELECTOR, "button.btn-primary"),
        (By.CSS_SELECTOR, "input[type='submit']"),
    ]
    for by, secici in yontemler:
        try:
            btn = driver.find_element(by, secici)
            driver.execute_script("arguments[0].click();", btn)
            print("   → Continue'ya basıldı ✓")
            return True
        except NoSuchElementException:
            continue
    return False


def a_input_bul(driver, wait):
    yontemler = [
        (By.XPATH, "//input[@placeholder='###-###-###']"),
        (By.XPATH, "//input[contains(@id,'aNumber')]"),
        (By.XPATH, "//input[contains(@name,'aNumber')]"),
        (By.XPATH, "//input[contains(@id,'alien')]"),
        (By.XPATH, "//input[contains(@placeholder,'###')]"),
        (By.CSS_SELECTOR, "input[type='text']"),
    ]
    for by, secici in yontemler:
        try:
            el = wait.until(EC.presence_of_element_located((by, secici)))
            return el
        except TimeoutException:
            continue
    return None


def hata_mesaji_al(driver):
    olasi_seciciler = [
        "//*[contains(text(), 'cannot be filed')]",
        "//*[contains(text(), 'A-Number not found')]",
        "//*[contains(text(), 'not found')]",
        "//div[contains(@class, 'error')]",
        "//ul[contains(@class, 'error')]",
    ]
    for secici in olasi_seciciler:
        try:
            el = driver.find_element(By.XPATH, secici)
            metin = el.text.strip()
            if metin:
                return metin
        except NoSuchElementException:
            continue
    return None


def main():
    print("=" * 50)
    print("  EOIR A-Number Kontrol Otomasyonu")
    print("  Yellow Law Group PC")
    print("=" * 50)
    print(f"\nExcel aranıyor: {EXCEL_DOSYASI}")

    try:
        wb = openpyxl.load_workbook(EXCEL_DOSYASI)
        ws = wb.active
    except FileNotFoundError:
        print(f"\nHATA: 'eoir_anumbers.xlsx' masaüstünde bulunamadı!")
        print("Lütfen Excel dosyasını masaüstüne kopyalayın.")
        input("\nÇıkmak için ENTER'a basın...")
        return

    if not ws.cell(row=1, column=SONUC_SUTUN).value:
        ws.cell(row=1, column=SONUC_SUTUN).value = "Sonuç"

    toplam = sum(1 for row in ws.iter_rows(min_row=BASLIK_SATIRI + 1)
                 if row[A_NUMBER_SUTUN - 1].value)
    print(f"Toplam {toplam} adet A-Number bulundu.\n")

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    wait = WebDriverWait(driver, BEKLEME_SURESI)

    driver.get(PORTAL_URL)
    time.sleep(2)
    agree_varsa_bas(driver)

    print("\n>>> Portal açıldı.")
    print(">>> Giriş yap (kullanıcı adı + şifre + 2FA).")
    print(">>> Giriş tamamlanınca buraya gel.\n")
    input("ENTER'a bas (devam etmek için)...")

    islenen = 0
    for row in ws.iter_rows(min_row=BASLIK_SATIRI + 1, max_row=ws.max_row):
        a_hucre = row[A_NUMBER_SUTUN - 1]
        b_hucre = row[SONUC_SUTUN - 1]
        if not a_hucre.value:
            continue

        a_number = a_numara_formatla(a_hucre.value)
        islenen += 1
        print(f"\n[{islenen}/{toplam}] İşleniyor: {a_number}")
        onceki_url = driver.current_url

        try:
            time.sleep(1)
            agree_varsa_bas(driver)
            a_input = a_input_bul(driver, wait)
            if not a_input:
                b_hucre.value = "HATA: A-Number kutusu bulunamadı"
                wb.save(EXCEL_DOSYASI); continue

            a_input.clear()
            time.sleep(0.3)
            a_input.send_keys(a_number)
            time.sleep(0.3)

            try:
                radio_buttons = driver.find_elements(By.XPATH, "//input[@type='radio']")
                if radio_buttons and not radio_buttons[0].is_selected():
                    radio_buttons[0].click()
            except Exception:
                pass

            basildimi = continue_butonuna_bas(driver)
            if not basildimi:
                b_hucre.value = "HATA: Continue butonu bulunamadı"
                wb.save(EXCEL_DOSYASI); continue

            time.sleep(SATIR_ARASI_BEKLEME)
            hata = hata_mesaji_al(driver)

            if hata:
                b_hucre.value = hata
                print(f"   → Hata: {hata[:60]}")
            elif driver.current_url != onceki_url:
                b_hucre.value = "Case is opened"
                print(f"   → ✅ Case is opened")
                driver.back()
                time.sleep(SATIR_ARASI_BEKLEME)
            else:
                sayfa_metni = driver.find_element(By.TAG_NAME, "body").text
                if "cannot" in sayfa_metni.lower() or "not found" in sayfa_metni.lower():
                    b_hucre.value = "A-Number not found"
                else:
                    b_hucre.value = "Belirsiz - manuel kontrol et"

        except TimeoutException:
            b_hucre.value = "HATA: Sayfa yüklenemedi"
        except Exception as e:
            b_hucre.value = f"HATA: {str(e)[:80]}"

        wb.save(EXCEL_DOSYASI)

    print(f"\n{'='*50}")
    print(f"✅ Bitti! {islened} A-Number işlendi.")
    print(f"Excel: {EXCEL_DOSYASI}")
    input("\nÇıkmak için ENTER'a basın...")
    driver.quit()


if __name__ == "__main__":
    main()
