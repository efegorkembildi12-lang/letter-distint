"""
parser.py — Dağıtım mektubu ve УПД PDF ayrıştırıcısı
Desteklenen belgeler:
  - Распределительное письмо (dağıtım mektubu)
  - УПД (Универсальный передаточный документ)
  - Счёт (fatura)
"""

import re
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Tuple, Union
from pathlib import Path

from io import StringIO
from pdfminer.high_level import extract_text as pdfminer_extract_text


# ─── Veri modelleri ────────────────────────────────────────────────────────────

@dataclass
class UPDItem:
    """УПД içindeki tek bir malzeme satırı"""
    material_name: str
    quantity: float
    unit: str
    price_excl_vat: float
    vat_rate: float           # 0.20 = %20
    total_incl_vat: float


@dataclass
class UPDDocument:
    """Ayrıştırılmış УПД belgesi"""
    number: str               # Örn: "51010880001"
    date: str                 # Örn: "29.03.2025"
    items: List[UPDItem] = field(default_factory=list)
    total_incl_vat: float = 0.0
    supplier_name: str = ""
    buyer_name: str = ""
    raw_text: str = ""        # Debug için


@dataclass
class DistributionLetter:
    """Ayrıştırılmış Распределительное письмо"""
    number: str               # Örn: "90"
    date: str                 # Örn: "31.03.2025"
    amount: float             # Örn: 17208850.0
    supplier_name: str        # Örn: "ООО «Элком-Электро»"
    supplier_inn: str         # Örn: "7703214111"
    invoice_number: str       # Örn: "00ЦБ-670951"
    invoice_date: str         # Örn: "04.03.2025"
    contract_number: str      # Örn: "ЭЭ/О-310822-4"
    specification_number: str # Örn: "12"
    upd_numbers: List[str] = field(default_factory=list)
    smeta_position: str = ""  # Örn: "2.1.1.3"
    raw_text: str = ""


@dataclass
class ParseResult:
    """Tek bir dosyanın ayrıştırma sonucu"""
    file_path: str
    doc_type: str             # "letter" | "upd" | "invoice" | "unknown"
    letter: Optional[DistributionLetter] = None
    upd: Optional[UPDDocument] = None
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ─── Yardımcı fonksiyonlar ─────────────────────────────────────────────────────

def _clean_amount(text: str) -> float:
    """
    "17 208 850,00" veya "17208850.00" → 17208850.0
    Rusça ondalık ayracı virgül, binlik ayracı boşluk
    """
    text = text.strip()
    # Binlik boşluk ve nbsp kaldır
    text = text.replace('\u00a0', '').replace(' ', '').replace('\u202f', '')
    # Virgülü noktaya çevir
    text = text.replace(',', '.')
    # Sembol temizle
    text = re.sub(r'[^\d.]', '', text)
    try:
        return float(text)
    except ValueError:
        return 0.0


def _extract_text_from_pdf(file_path: str) -> str:
    """PDF'den ham metin çıkar (pdfminer.six)"""
    text = pdfminer_extract_text(file_path)
    return text or ""
# ─── Belge tipi tespiti ────────────────────────────────────────────────────────

def detect_doc_type(text: str) -> str:
    """
    PDF metni inceleyerek belge tipini tespit et.
    Dönüş: "letter" | "upd" | "invoice" | "spec" | "contract" | "unknown"
    Önce kesin belirleyicileri (strong markers) kontrol et, sonra puanlama yap.
    """
    text_lower = text.lower()

    # Kesin belirleyiciler — bunlardan biri varsa direkt karar ver
    strong = {
        "letter":  ["распределительное письмо", "просим вас оплатить поставщику",
                    "просим вас произвести оплату", "распорядительного письма"],
        "upd":     ["универсальный передаточный документ"],
        "invoice": ["счет на оплату №", "счёт на оплату №"],
        "spec":    ["спецификация №", "поставщик поставляет товар"],
        "contract":["договор поставки №", "договор строительного подряда"],
    }
    for doc_type, terms in strong.items():
        for term in terms:
            if term in text_lower:
                return doc_type

    # Zayıf puanlama — strong eşleşme yoksa
    weak = {
        "upd":     ["грузоотправитель", "грузополучатель",
                    "счет-фактура", "передаточный"],
        "invoice": ["счет №", "счёт №", "итого с ндс"],
        "letter":  ["исх.№", "просим"],
    }
    scores = {k: 0 for k in weak}
    for doc_type, terms in weak.items():
        for term in terms:
            if term in text_lower:
                scores[doc_type] += 1
    best = max(scores, key=scores.get)
    return best if scores[best] >= 2 else "unknown"


# ─── Дağıtım mektubu ayrıştırıcı ──────────────────────────────────────────────

def parse_distribution_letter(text: str, file_path: str = "") -> ParseResult:
    """Распределительное письмо'yu ayrıştır"""
    result = ParseResult(file_path=file_path, doc_type="letter")
    errors = []

    # Mektup numarası ve tarihi
    # Örn: "исх.№90 от 31.03.2025" veya "исх. № 90 от 31.03.2025"
    num_match = re.search(
        r'исх[.\s]*№\s*(\d+)\s+от\s+(\d{2}\.\d{2}\.\d{4})',
        text, re.IGNORECASE
    )
    letter_number = num_match.group(1) if num_match else ""
    letter_date   = num_match.group(2) if num_match else ""
    if not num_match:
        errors.append("Mektup numarası/tarihi bulunamadı")

    # Ödeme tutarı — birden fazla format
    amount_match = re.search(
        r'сумм[еу][:\s]*([\d\s\u00a0]+[,.]?\d*)\s*(?:руб|₽)',
        text, re.IGNORECASE
    )
    if not amount_match:
        # Alternatif: büyük sayı ara
        amount_match = re.search(
            r'((?:\d{1,3}[\s\u00a0])*\d{1,3}[,]\d{2})\s*(?:руб|₽)',
            text
        )
    amount = _clean_amount(amount_match.group(1)) if amount_match else 0.0
    if amount == 0.0:
        errors.append("Ödeme tutarı bulunamadı veya sıfır")

    # Tedarikçi: satır atlamasını destekle
    # Gerçek format: satır1="поставщику\nОбщество...«Элком-Электро» (ООО «Элком-Электро»)"
    supplier_name = ""
    # "оплатить поставщику" + sonraki satırda veya aynı satırda «...» (ООО «...»)
    pay_ctx = re.search(
        r'оплатить\s+поставщику[\s\S]{0,200}?\((ООО|ОАО|ЗАО|АО|ПАО)\s*[«"]([^»"\n]{3,60})[»"]\)',
        text, re.IGNORECASE
    )
    if pay_ctx:
        supplier_name = f"{pay_ctx.group(1)} «{pay_ctx.group(2)}»"
    else:
        # Fallback: «...» (ООО «...») pattern — kısa isim parantez içinde
        short_name = re.search(
            r'\(((?:ООО|ОАО|ЗАО|АО|ПАО)\s*[«"]([^»"\n]{3,40})[»"])\)',
            text
        )
        if short_name:
            supplier_name = short_name.group(1).strip()
        else:
            sm = re.search(r'((?:ООО|ОАО|ЗАО|АО|ПАО)\s*[«"]([^»"\n]{3,40})[»"])', text)
            if sm:
                supplier_name = sm.group(1).strip()

    # Tedarikçi INN — "ИНН/КПП 7703214111/771701001" veya "ИНН 7703214111"
    # Supplier bağlamında INN bul (mektup sahibinin değil, tedarikçinin)
    inn_ctx = re.search(
        r'(?:Элком|поставщику|оплатить)[^ИНН]{0,300}ИНН[/\s]*[:\s]?(\d{10})',
        text, re.DOTALL | re.IGNORECASE
    )
    if inn_ctx:
        supplier_inn = inn_ctx.group(1)
    else:
        inn_match = re.search(r'ИНН[/]?КПП\s+(\d{10})', text)
        if not inn_match:
            inn_match = re.search(r'ИНН\s*[:\s]?(\d{10})', text)
        supplier_inn = inn_match.group(1) if inn_match else ""

    # Fatura numarası
    # "сч.№ЦБ-670951" veya "счет №00ЦБ-670951"
    # Счёт: "счет на оплату № 00ЦБ-670951 от 04.03.2025" veya "сч.№ЦБ-670951"
    inv_match = re.search(
        r'(?:[сС][чЧ][её][тТ][а-я]*\s+(?:на\s+оплату\s+)?№\s*|[сС][чЧ]\.\s*№\s*)([\w\-]+)\s+от\s+(\d{2}\.\d{2}\.\d{4})',
        text
    )
    invoice_number = inv_match.group(1).strip() if inv_match else ""
    invoice_date   = inv_match.group(2) if inv_match else ""

    # Sözleşme numarası — Договор поставки № ЭЭ/О-310822-4
    # Önce поставки sözleşmesini ara
    contract_match = re.search(
        r'[Дд]оговор[а-я\s]*поставки\s*№\s*([\w\-\/\.]+)',
        text
    )
    if not contract_match:
        contract_match = re.search(r'[Дд]оговор[а-я\s]*№\s*([\w\-\/\.]+)', text)
    contract_number = contract_match.group(1) if contract_match else ""

    # Spesifikasyon numarası
    spec_match = re.search(
        r'[Сс]пецификаци[яи][^\n]*№\s*(\d+)',
        text
    )
    spec_number = spec_match.group(1) if spec_match else ""

    # УПД numaraları — "УПД ООО «...» № 51010880001" (satır içi)
    upd_numbers = re.findall(r'УПД[^\d\n]{0,50}№\s*(\d{8,12})', text, re.IGNORECASE)
    if not upd_numbers:
        # Fallback: 11 haneli sayılar (УПД numarası formatı)
        upd_numbers = re.findall(r'№\s*(\d{11})', text)

    # Smeta pozisyon referansı — "2.1.1.3" (Сметы sütununda)
    smeta_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', text)
    smeta_pos = smeta_match.group(1) if smeta_match else ""

    result.letter = DistributionLetter(
        number=letter_number,
        date=letter_date,
        amount=amount,
        supplier_name=supplier_name,
        supplier_inn=supplier_inn,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        contract_number=contract_number,
        specification_number=spec_number,
        upd_numbers=upd_numbers,
        smeta_position=smeta_pos,
        raw_text=text,
    )
    result.errors = errors
    return result


# ─── УПД ayrıştırıcı ──────────────────────────────────────────────────────────

def parse_upd(text: str, file_path: str = "") -> ParseResult:
    """УПД belgesini ayrıştır"""
    result = ParseResult(file_path=file_path, doc_type="upd")

    # УПД numarası ve tarihi
    num_match = re.search(
        r'(?:УПД|№)\s*(?:№\s*)?(\d{8,12})\s+от\s+(\d{2}\.\d{2}\.\d{4})',
        text, re.IGNORECASE
    )
    upd_number = num_match.group(1) if num_match else ""
    upd_date   = num_match.group(2) if num_match else ""

    # Alternatif: dosya adından УПД numarası çek
    if not upd_number and file_path:
        fn_match = re.search(r'(\d{8,12})', Path(file_path).stem)
        if fn_match:
            upd_number = fn_match.group(1)

    # Toplam tutar
    total_match = re.search(
        r'[Ии]того\s+(?:с\s+НДС\s+)?(?:[\d\s\u00a0,]+)\s*([\d\s\u00a0]+[,]\d{2})',
        text
    )
    if not total_match:
        # Alternatif pattern
        total_match = re.search(
            r'(?:Сумма\s+с\s+НДС|Итого)[:\s]*([\d\s\u00a0]+[,]\d{2})',
            text, re.IGNORECASE
        )
    total = _clean_amount(total_match.group(1)) if total_match else 0.0

    # Tedarikçi (Продавец satırı)
    supplier_match = re.search(r'[Пп]родавец[:\s]+([^\n]+)', text)
    supplier = supplier_match.group(1).strip() if supplier_match else ""

    # Alıcı (Покупатель satırı)
    buyer_match = re.search(r'[Пп]окупатель[:\s]+([^\n]+)', text)
    buyer = buyer_match.group(1).strip() if buyer_match else ""

    # Malzeme satırlarını çıkarmak için tablo ayrıştırma
    items = _extract_upd_items(text)

    result.upd = UPDDocument(
        number=upd_number,
        date=upd_date,
        items=items,
        total_incl_vat=total,
        supplier_name=supplier,
        buyer_name=buyer,
        raw_text=text,
    )
    return result


def _extract_upd_items(text: str) -> List[UPDItem]:
    """
    УПД tablo satırlarını ayrıştır.
    Kablo ürünleri için optimize edilmiş.
    """
    items = []

    # Kablo adı pattern: "Кабель [МАРКА] [СЕЧЕНИЕ]"
    cable_pattern = re.compile(
        r'((?:Кабель|КАБЕЛЬ)\s+[\w\(\)А-Яа-яёЁ\-\s]+?'
        r'(?:\d+[хx\*×]\d+(?:[,\.]\d+)?(?:\s*(?:мк-\d|мм²?|кВ))?)'
        r'[^\n]*?)'
        r'\s+([\d\s\u00a0]+[,]\d+)\s+'  # количество
        r'(м|шт|компл)\s+'               # ед.изм
        r'([\d\s\u00a0]+[,]\d+)\s+'      # цена без НДС
        r'(\d+)\s*%?',                    # НДС %
        re.IGNORECASE
    )

    for m in cable_pattern.finditer(text):
        name = m.group(1).strip()
        qty  = _clean_amount(m.group(2))
        unit = m.group(3)
        price = _clean_amount(m.group(4))
        vat   = float(m.group(5)) / 100
        total = round(qty * price * (1 + vat), 2)
        items.append(UPDItem(
            material_name=name,
            quantity=qty,
            unit=unit,
            price_excl_vat=price,
            vat_rate=vat,
            total_incl_vat=total,
        ))

    return items


# ─── Ana giriş noktası ─────────────────────────────────────────────────────────

def parse_pdf(file_path: str) -> ParseResult:
    """
    Tek bir PDF dosyasını ayrıştır.
    Belge tipini otomatik tespit eder.
    """
    path = Path(file_path)
    if not path.exists():
        return ParseResult(
            file_path=file_path,
            doc_type="unknown",
            errors=[f"Dosya bulunamadı: {file_path}"]
        )
    if path.suffix.lower() not in (".pdf", ".PDF"):
        return ParseResult(
            file_path=file_path,
            doc_type="unknown",
            errors=["Sadece PDF dosyaları destekleniyor"]
        )

    try:
        text = _extract_text_from_pdf(file_path)
    except ImportError as e:
        return ParseResult(
            file_path=file_path,
            doc_type="unknown",
            errors=[str(e)]
        )
    except Exception as e:
        return ParseResult(
            file_path=file_path,
            doc_type="unknown",
            errors=[f"PDF okunamadı: {str(e)}"]
        )

    doc_type = detect_doc_type(text)

    if doc_type == "letter":
        return parse_distribution_letter(text, file_path)
    elif doc_type == "upd":
        return parse_upd(text, file_path)
    else:
        return ParseResult(
            file_path=file_path,
            doc_type=doc_type,
            warnings=[f"Belge tipi tespit edilemedi ('{doc_type}') — manuel kontrol gerekli"]
        )


def parse_batch(file_paths: List[str]) -> List[ParseResult]:
    """Birden fazla PDF dosyasını toplu ayrıştır"""
    return [parse_pdf(fp) for fp in file_paths]
