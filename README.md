<div align="center">

# letter-distint

**Automate Russian business PDF → Excel data entry. No manual copy-paste.**

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![Languages](https://img.shields.io/badge/languages-TR%20%7C%20RU%20%7C%20EN-orange)
![License](https://img.shields.io/badge/license-MIT-green)

Reads Russian supply chain PDF documents (distribution letters, UPDs, invoices),
validates amounts, matches materials to your Excel template, and writes the data — all in a few clicks.

[English](#english) · [Русский](#русский)

</div>

---

<a name="english"></a>

## What it does

Russian procurement workflows involve three document types that must be reconciled and entered into Excel workbooks manually. **letter-distint** automates this:

1. **Parse** — extracts supplier names, invoice numbers, amounts, and material lists from PDF files
2. **Validate** — checks that UPD totals match the distribution letter amount (highlights mismatches)
3. **Match** — finds the correct supplier columns in your Excel workbook automatically
4. **Write** — fills in the data; original file is never overwritten (`_updated` copy is created)

A dry-run preview shows exactly what will be written before any file is touched.

---

## Supported document types

| Document | Russian name | Description |
|---|---|---|
| Distribution Letter | Распределительное письмо | Master document with total amount and supplier breakdown |
| UPD | Универсальный передаточный документ | Combined invoice + delivery document |
| Invoice | Счёт | Standard supplier invoice |

---

## Who is this for?

- Finance and procurement teams working with **Russian suppliers**
- Anyone who manually copies data from Russian PDF documents into spreadsheets

If you receive Russian distribution letters, match them against UPDs, and fill in Excel — this tool eliminates that work.

---

## Download (Windows)

Download the latest `letter-distint.exe` from the [Releases](../../releases) page and run it directly. No Python installation required.

---

## Run from source

```bash
git clone https://github.com/efegorkembildi12-lang/letter-distint.git
cd letter-distint
pip install -r requirements.txt
python app.py
```

**Requirements:** Python 3.8+, `pdfminer.six`, `openpyxl >= 3.1.0`

---

## Usage

The app has four tabs:

| Tab | What you do |
|---|---|
| **1 · Files** | Add PDF files (letter + UPDs), select your Excel workbook and sheet |
| **2 · Preview** | Process PDFs — see extracted data and amount validation |
| **3 · Excel Write** | Match data to columns, preview the write plan (dry-run), then write |
| **4 · Log** | Full operation log for review or debugging |

Use the **TR / RU / EN** button in the top-right corner to switch the interface language.

---

## Build from source (Windows .exe)

```bash
pip install pyinstaller
pyinstaller letter-distint.spec
# Output: dist/letter-distint.exe
```

---

## Excel template requirements

Your Excel workbook must contain:
- A material list with supplier columns (volume, price, cost, payment)
- Sheet name is specified in the app (Tab 1)

---

## License

MIT © [Efe Görkem Bildi](https://github.com/efegorkembildi12-lang)

---

<a name="русский"></a>

<div align="center">

## Русский

**Автоматизация переноса данных из российских PDF-документов в Excel. Без ручного копирования.**

</div>

---

## Что делает программа

При работе с российскими поставщиками приходится вручную сверять три типа документов и переносить данные в Excel-файлы. **letter-distint** автоматизирует этот процесс:

1. **Парсинг** — извлекает наименования поставщиков, номера счетов, суммы и перечень материалов из PDF
2. **Проверка** — сверяет итоговые суммы УПД с суммой из распределительного письма (расхождения выделяются)
3. **Сопоставление** — автоматически находит нужные столбцы поставщиков в Excel-файле
4. **Запись** — вносит данные в файл; оригинал не изменяется (создаётся копия с суффиксом `_updated`)

Перед записью можно просмотреть план изменений в режиме dry-run.

---

## Поддерживаемые типы документов

| Документ | Описание |
|---|---|
| Распределительное письмо | Основной документ с итоговой суммой и разбивкой по поставщикам |
| УПД (универсальный передаточный документ) | Совмещённый счёт-фактура и накладная |
| Счёт | Стандартный счёт поставщика |

---

## Для кого предназначена программа

- Финансовые и закупочные отделы, работающие с **российскими поставщиками**
- Все, кто вручную переносит данные из российских PDF-документов в Excel

Если вы получаете распределительные письма, сверяете их с УПД и заполняете Excel — эта программа избавит вас от этой работы.

---

## Скачать (Windows)

Скачайте `letter-distint.exe` со страницы [Releases](../../releases) и запустите напрямую. Установка Python не требуется.

---

## Запуск из исходного кода

```bash
git clone https://github.com/efegorkembildi12-lang/letter-distint.git
cd letter-distint
pip install -r requirements.txt
python app.py
```

**Требования:** Python 3.8+, `pdfminer.six`, `openpyxl >= 3.1.0`

---

## Использование

Приложение состоит из четырёх вкладок:

| Вкладка | Действие |
|---|---|
| **1 · Файлы** | Добавьте PDF-файлы (письмо + УПД), выберите Excel-файл и лист |
| **2 · Предпросмотр** | Обработайте PDF — просмотрите извлечённые данные и результат сверки |
| **3 · Запись в Excel** | Сопоставьте данные со столбцами, просмотрите план записи (dry-run), затем запишите |
| **4 · Журнал** | Полный журнал операций для просмотра и отладки |

Кнопка **TR / RU / EN** в правом верхнем углу переключает язык интерфейса.

---

## Сборка .exe из исходного кода

```bash
pip install pyinstaller
pyinstaller letter-distint.spec
# Результат: dist/letter-distint.exe
```

---

## Требования к шаблону Excel

Excel-файл должен содержать:
- Перечень материалов со столбцами поставщиков (объём, цена, стоимость, оплата)
- Имя листа указывается в приложении (Вкладка 1)

---

## Лицензия

MIT © [Efe Görkem Bildi](https://github.com/efegorkembildi12-lang)
