#!/bin/bash
GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'

DESKTOP="$HOME/Desktop"
TOOLS_DIR="$DESKTOP/YellowLawTools"

clear
echo "=============================================="
echo "   Yellow Law Group PC - Araç Kurulumu"
echo "=============================================="
echo ""
echo "Hangi araçları kurmak istiyorsunuz?"
echo ""
echo "  1) EOIR A-Number Sorgulama"
echo "  2) USCIS Case Status Sorgulama"
echo "  3) Annual Asylum Fee Kontrolü"
echo "  4) Hepsini Kur"
echo ""
read -p "Seçiminiz (örn: 1 veya 1,3 veya 4): " SECIM
echo ""

# Python kontrolü
echo -e "${YELLOW}Python kontrol ediliyor...${NC}"
if ! command -v python3 &>/dev/null; then
    echo -e "${RED}Python bulunamadı!${NC}"
    echo "Lütfen python.org/downloads adresinden Python'ı yükleyin."
    read -p "ENTER'a basın..."
    exit 1
fi
echo -e "${GREEN}✅ Python mevcut${NC}"

# Paketleri kur
echo -e "${YELLOW}Gerekli paketler yükleniyor...${NC}"
pip3 install selenium openpyxl webdriver-manager playwright --quiet --break-system-packages 2>/dev/null || \
pip3 install selenium openpyxl webdriver-manager playwright --quiet 2>/dev/null
python3 -m playwright install chromium --quiet 2>/dev/null
echo -e "${GREEN}✅ Paketler yüklendi${NC}"

# Tools klasörü
mkdir -p "$TOOLS_DIR"

# Scripti indir ve .command dosyası oluştur
install_tool() {
    local file=$1
    local display=$2
    local excel_adi=$3

    echo -e "${YELLOW}$display kuruluyor...${NC}"

    # GitHub'dan indir
    curl -fsSL "https://raw.githubusercontent.com/corpuso233/yellowlawtools/main/tools/$file" \
         -o "$TOOLS_DIR/$file"

    if [ ! -s "$TOOLS_DIR/$file" ]; then
        echo -e "${RED}İndirme başarısız! İnternet bağlantınızı kontrol edin.${NC}"
        return
    fi

    # Excel şablonu oluştur
    python3 -c "
import openpyxl, os
wb = openpyxl.Workbook()
ws = wb.active
headers = '''$excel_adi'''.split('|')
for i, h in enumerate(headers, 1):
    ws.cell(row=1, column=i).value = h.strip()
path = os.path.join(os.path.expanduser('~'), 'Desktop', headers[0].strip().lower().replace(' ','-').replace('/','-') + '-template.xlsx')
# Use fixed names
" 2>/dev/null

    # .command dosyası (çift tıkla çalışır)
    local cmd_file="$TOOLS_DIR/${display// /_}.command"
    cat > "$cmd_file" << CMD
#!/bin/bash
cd "$TOOLS_DIR"
python3 "$file"
CMD
    chmod +x "$cmd_file"
    echo -e "${GREEN}✅ $display kuruldu → ${display// /_}.command${NC}"
}

if [[ "$SECIM" == *"1"* ]] || [[ "$SECIM" == "4" ]]; then
    install_tool "eoir_otomasyon.py" "EOIR_A-Number_Sorgulama" "A-Number|Sonuç"

    # Excel şablonu
    python3 -c "
import openpyxl, os
wb = openpyxl.Workbook(); ws = wb.active
ws['A1'] = 'A-Number'; ws['B1'] = 'Sonuç'
wb.save(os.path.join(os.path.expanduser('~'),'Desktop','eoir_anumbers.xlsx'))
print('Excel şablonu: eoir_anumbers.xlsx')
"
fi

if [[ "$SECIM" == *"2"* ]] || [[ "$SECIM" == "4" ]]; then
    install_tool "uscis_case_status.py" "USCIS_Case_Status" "Receipt Number|Case Status|Annual Fee Notu|Açıklama"

    python3 -c "
import openpyxl, os
wb = openpyxl.Workbook(); ws = wb.active
ws['A1'] = 'Receipt Number'; ws['B1'] = 'Case Status'; ws['C1'] = 'Annual Fee Notu'; ws['D1'] = 'Açıklama'
wb.save(os.path.join(os.path.expanduser('~'),'Desktop','uscis_receipt_numbers.xlsx'))
print('Excel şablonu: uscis_receipt_numbers.xlsx')
"
fi

if [[ "$SECIM" == *"3"* ]] || [[ "$SECIM" == "4" ]]; then
    install_tool "annual_fee_check.py" "Annual_Fee_Kontrolu" "A-Number|Receipt Number|Sonuç"

    python3 -c "
import openpyxl, os
wb = openpyxl.Workbook(); ws = wb.active
ws['A1'] = 'A-Number'; ws['B1'] = 'Receipt Number'; ws['C1'] = 'Sonuç'
wb.save(os.path.join(os.path.expanduser('~'),'Desktop','annual_fee.xlsx'))
print('Excel şablonu: annual_fee.xlsx')
"
fi

echo ""
echo "=============================================="
echo -e "${GREEN}✅ Kurulum tamamlandı!${NC}"
echo ""
echo "Masaüstündeki 'YellowLawTools' klasörüne gidin."
echo "İlgili .command dosyasına çift tıklayın."
echo ""
echo "Excel şablonları masaüstünüze kaydedildi."
echo "Verileri girin ve aracı çalıştırın."
echo "=============================================="
