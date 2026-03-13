"""
test_parser.py — Parser modülü birim testleri
Çalıştır: python -m pytest tests/ -v
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.parser import (
    detect_doc_type,
    parse_distribution_letter,
    _clean_amount,
)
from core.excel_writer import extract_invoice_key as excel_extract_key


# ─── _clean_amount testleri ────────────────────────────────────────────────────

def test_clean_amount_russian_format():
    """Rusça format: binlik boşluk + virgüllü ondalık"""
    assert _clean_amount("17 208 850,00") == 17208850.0
    assert _clean_amount("133 650,00") == 133650.0
    assert _clean_amount("1 533,33") == 1533.33

def test_clean_amount_dot_decimal():
    assert _clean_amount("17208850.00") == 17208850.0

def test_clean_amount_nbsp():
    """Non-breaking space ile binlik ayraç"""
    assert _clean_amount("17\u00a0208\u00a0850,00") == 17208850.0

def test_clean_amount_zero():
    assert _clean_amount("0,00") == 0.0
    assert _clean_amount("") == 0.0


# ─── detect_doc_type testleri ──────────────────────────────────────────────────

def test_detect_letter():
    text = """
    ООО «ЮК Инжиниринг»
    Распределительное письмо исх.№90 от 31.03.2025
    Просим вас произвести оплату в адрес ООО «Элком-Электро»
    """
    assert detect_doc_type(text) == "letter"

def test_detect_upd():
    text = """
    Универсальный передаточный документ №51010880001 от 29.03.2025
    Продавец: ООО «Элком-Электро»
    Покупатель: ООО «ЮК Инжиниринг»
    Грузоотправитель: склад
    """
    assert detect_doc_type(text) == "upd"

def test_detect_invoice():
    text = """
    Счёт на оплату №00ЦБ-670951 от 04.03.2025
    ООО «Элком-Электро»
    """
    assert detect_doc_type(text) == "invoice"

def test_detect_unknown():
    assert detect_doc_type("Произвольный текст без ключевых слов") == "unknown"


# ─── parse_distribution_letter testleri ───────────────────────────────────────

SAMPLE_LETTER_TEXT = """
ООО «Смайнэкс Констракшн»

Исх.№90 от 31.03.2025

Распределительное письмо

Просим вас произвести оплату в адрес ООО «Элком-Электро» (ИНН 7703214111)
по счёт №00ЦБ-670951 от 04.03.2025 на сумму: 17 208 850,00 руб.

Основание: Спецификация №12 от 04.03.2025 к Договору поставки №ЭЭ/О-310822-4
от 31.08.2022г.

Материалы согласно п.2.1.1.3:
- УПД №51010880001
- УПД №52010880049

С уважением,
ООО «ЮК Инжиниринг»
"""

def test_parse_letter_number():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert result.letter is not None
    assert result.letter.number == "90"
    assert result.letter.date == "31.03.2025"

def test_parse_letter_amount():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert result.letter.amount == 17208850.0

def test_parse_letter_supplier():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert "Элком-Электро" in result.letter.supplier_name
    assert result.letter.supplier_inn == "7703214111"

def test_parse_letter_invoice():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert "670951" in result.letter.invoice_number
    assert result.letter.invoice_date == "04.03.2025"

def test_parse_letter_contract():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert "ЭЭ/О-310822-4" in result.letter.contract_number

def test_parse_letter_specification():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert result.letter.specification_number == "12"

def test_parse_letter_upd_numbers():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert "51010880001" in result.letter.upd_numbers
    assert "52010880049" in result.letter.upd_numbers

def test_parse_letter_smeta_position():
    result = parse_distribution_letter(SAMPLE_LETTER_TEXT)
    assert result.letter.smeta_position == "2.1.1.3"


# ─── extract_invoice_key testleri ─────────────────────────────────────────────

def test_extract_key_prefix():
    """"00ЦБ-670951" → "670951" """
    assert excel_extract_key("00ЦБ-670951") == "670951"

def test_extract_key_slash():
    """"969/2504067-1" → "2504067" (en uzun sayısal blok) """
    key = excel_extract_key("969/2504067-1")
    assert key == "2504067"

def test_extract_key_plain():
    assert excel_extract_key("670951") == "670951"


# ─── Çalıştırma ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import traceback
    tests = [
        test_clean_amount_russian_format,
        test_clean_amount_dot_decimal,
        test_clean_amount_nbsp,
        test_clean_amount_zero,
        test_detect_letter,
        test_detect_upd,
        test_detect_invoice,
        test_detect_unknown,
        test_parse_letter_number,
        test_parse_letter_amount,
        test_parse_letter_supplier,
        test_parse_letter_invoice,
        test_parse_letter_contract,
        test_parse_letter_specification,
        test_parse_letter_upd_numbers,
        test_parse_letter_smeta_position,
        test_extract_key_prefix,
        test_extract_key_slash,
        test_extract_key_plain,
    ]
    passed = failed = 0
    for t in tests:
        try:
            t()
            print(f"  ✓  {t.__name__}")
            passed += 1
        except Exception as e:
            print(f"  ✗  {t.__name__}: {e}")
            failed += 1
    print(f"\n{'─'*40}")
    print(f"  {passed} geçti · {failed} başarısız")
